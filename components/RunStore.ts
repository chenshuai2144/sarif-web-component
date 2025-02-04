/* eslint-disable @typescript-eslint/no-parameter-properties */
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { observable, autorun, computed, IObservableValue } from 'mobx';
import { Run, Result, Artifact } from 'sarif';

import { MobxFilter } from './FilterBar';
import { tryOr } from './try';
import { Rule, ResultOrRuleOrMore } from './Viewer.Types';

import { SortOrder } from 'azure-devops-ui/Table';
import { ITreeItem } from 'azure-devops-ui/Utilities/TreeItemProvider';

declare module 'sarif' {
  interface Run {
    _augmented: boolean;
    _rulesInUse: Map<string, Rule>;
    _agesInUse: Map<
      string,
      {
        results: any[];
        treeItem: any;
        name: string;
        isAge: boolean;
      }
    >;
  }
}

export enum SortRuleBy {
  Count,
  Name,
}

export const isMatch = (field: string, keywords: string[]) =>
  !keywords.length || keywords.some((keyword) => field.includes(keyword));

export class RunStore {
  driverName: string;
  @observable sortRuleBy = SortRuleBy.Count;
  @observable sortColumnIndex = 1;
  @observable sortOrder = SortOrder.ascending;

  constructor(
    readonly run: Run,
    readonly logIndex,
    readonly filter: MobxFilter,
    readonly groupByAge?: IObservableValue<boolean>,
    readonly hideBaseline?: boolean,
    readonly showAge?: boolean,
    readonly hideLevel?: boolean,
  ) {
    const { driver } = run.tool;
    this.driverName =
      (run.properties && run.properties['logFileName']) ||
      (driver.fullName || driver.name).replace(
        /^Microsoft.CodeAnalysis.Sarif.PatternMatcher$/,
        'CredScan on Push',
      );

    if (!run._augmented) {
      run._rulesInUse = new Map<string, Rule>();
      run._agesInUse = new Map([
        ['Past SLA', { results: [], treeItem: null, name: 'Past SLA (31+ days)', isAge: true }],
        [
          'Within SLA',
          { results: [], treeItem: null, name: 'Within SLA (0 - 30 days)', isAge: true },
        ],
      ]);

      const rules = driver.rules || [];
      const rulesListed = new Map<string, Rule>(rules.map((rule) => [rule.id, rule] as any)); // Unable to express [[string, RuleEx]].
      run.results?.forEach((result) => {
        // Collate by Rule
        const { ruleIndex } = result;
        const ruleId = result.ruleId ?? '(No Rule)'; // Ignores 3.5.4 Hierarchical strings.
        if (!run._rulesInUse.has(ruleId)) {
          // Need to generate rules for some like Microsoft.CodeAnalysis.Sarif.PatternMatcher.
          const rule =
            (ruleIndex !== undefined && (rules[ruleIndex] as Rule)) ||
            rulesListed.get(ruleId) ||
            ({ id: ruleId } as Rule);
          rule.isRule = true;
          rule.run = run; // For taxa.
          run._rulesInUse.set(ruleId, rule);
        }

        const rule = run._rulesInUse.get(ruleId);
        rule.results = rule.results || [];
        rule.results.push(result);

        // Collate by Age
        const firstDetection = result.provenance?.firstDetectionTimeUtc;
        result.firstDetection = firstDetection ? new Date(firstDetection) : new Date();
        const age =
          (new Date().getTime() - result.firstDetection.getTime()) / (24 * 60 * 60 * 1000); // 1 day in milliseconds
        result.sla = age > 31 ? 'Past SLA' : 'Within SLA';
        run._agesInUse.get(result.sla).results.push(result);

        // Fill-in url from run.artifacts as needed.
        const artLoc = tryOr(() => result.locations[0].physicalLocation.artifactLocation);
        if (artLoc && artLoc.uri === undefined) {
          const art = tryOr<Artifact>(() => run.artifacts[artLoc.index]);
          artLoc.uri = art.location?.uri;
        }

        result.run = run; // For result renderer to get to run.artifacts.
        result._rule = rule;
      });
      run._augmented = true;
    }

    autorun(
      () => {
        this.showAllRevision; // Read.
        const rules = this.groupByAge.get() // Slice to satisfy ref rulesTruncated.
          ? this.agesFiltered.slice()
          : this.rulesFiltered.slice();

        rules.forEach((ruleTreeItem) => {
          const maxLength = 3;
          ruleTreeItem.childItems =
            !ruleTreeItem.isShowAll && ruleTreeItem.childItemsAll.length > maxLength
              ? [
                  ...ruleTreeItem.childItemsAll.slice(0, maxLength),
                  {
                    data: {
                      onClick: () => {
                        ruleTreeItem.isShowAll = true;
                        this.showAllRevision++;
                      },
                    },
                  },
                ]
              : ruleTreeItem.childItemsAll;
        });

        this.rulesTruncated = rules;
      },
      { name: 'Truncation' },
    );
  }

  private filterHelper(treeItems: ITreeItem<ResultOrRuleOrMore>[]) {
    const filter = this.filter.getState();
    const filterKeywords = (filter.Keywords?.value ?? '')
      .toLowerCase()
      .split(/\s+/)
      .filter((part) => part);
    const { sortColumnIndex, sortOrder } = this;

    treeItems.forEach((treeItem) => {
      // if (!treeItem.hasOwnProperty('isShowAll')) extendObservable(treeItem, { isShowAll: false })
      treeItem.isShowAll = false;

      // Filtering logic: Show if 1) dropdowns match AND 2) any field matches text.
      const isDriverMatch = isMatch(this.driverName.toLowerCase(), filterKeywords);

      const resultContainer = treeItem.data as { results: Result[] };
      treeItem.childItemsAll = resultContainer.results
        .filter((result) => {
          const { _rule } = result;
          const ruleId = _rule.id.toLowerCase();
          const ruleName = _rule.name?.toLowerCase() ?? '';
          const isRuleMatch = isMatch(ruleId, filterKeywords) || isMatch(ruleName, filterKeywords);

          for (const columnName in filter) {
            if (columnName === 'Discussion') continue; // Discussion filter does not apply to Results.
            const selectedValues = filter[columnName].value;
            if (!Array.isArray(selectedValues)) continue;
            if (!selectedValues.length) continue;
            const map = {
              Baseline: (result: Result) => (result.baselineState as string) || 'new', // TODO: Merge with column def.
              Level: (result: Result) => result.level || 'warning',
              Suppression: (result: Result) =>
                result.suppressions?.some((s) => s.status === undefined || s.status === 'accepted')
                  ? 'suppressed'
                  : 'unsuppressed',
              Age: (result: Result) => result.sla.toLowerCase(),
            };
            const translatedCellValue = map[columnName] ? map[columnName](result) : result;
            if (!selectedValues.includes(translatedCellValue)) return false;
          }

          const isKeywordMatch = this.columns.some((column) => {
            const field = column.filterString(result).toLowerCase();
            return isMatch(field, filterKeywords);
          });

          return isDriverMatch || isRuleMatch || isKeywordMatch;
        })
        .map((result) => ({ data: result })); // Can cache the result here.

      treeItem.childItemsAll.sort((treeItemLeft, treeItemRight) => {
        const resultToValue = this.columns[sortColumnIndex].sortString;
        const valueLeft = resultToValue(treeItemLeft.data as Result);
        const valueRight = resultToValue(treeItemRight.data as Result);

        const inverter = sortOrder === SortOrder.ascending ? 1 : -1;
        return inverter * valueLeft.localeCompare(valueRight);
      });

      return treeItem as ITreeItem<ResultOrRuleOrMore>;
    });

    const treeItemsVisible = treeItems.filter((rule) => rule.childItemsAll.length);

    treeItemsVisible.sort(
      this.sortRuleBy === SortRuleBy.Count
        ? (a, b) => b.childItemsAll.length - a.childItemsAll.length
        : (a, b) => (a.data as Rule).id.localeCompare((b.data as Rule).id),
    );

    treeItemsVisible.forEach((rule, i) => (rule.expanded = i === 0));

    return treeItemsVisible;
  }

  @computed get agesFiltered() {
    const treeItems = [...this.run._agesInUse.values()].map((age) => {
      const treeItem = (age.treeItem = age.treeItem || {
        data: age,
        expanded: false,
      });
      return treeItem as ITreeItem<ResultOrRuleOrMore>;
    });
    return this.filterHelper(treeItems);
  }

  @computed get rulesFiltered() {
    const treeItems = [...this.run._rulesInUse.values()].map((rule) => {
      const treeItem = (rule.treeItem = rule.treeItem || {
        data: rule,
        expanded: false,
      });
      return treeItem as ITreeItem<ResultOrRuleOrMore>;
    });
    return this.filterHelper(treeItems);
  }

  @computed get filteredCount() {
    return this.rulesFiltered.reduce((total, rule) => total + rule.childItemsAll.length, 0);
  }

  @observable showAllRevision = 0;
  @observable.ref rulesTruncated = [] as ITreeItem<ResultOrRuleOrMore>[]; // Technically ITreeItem<Rule>[], ref assuming immutable array.

  @computed get columns() {
    const columns = [
      {
        id: 'Path',
        text: '漏洞名称',
        filterString: (result: Result) =>
          tryOr<string>(
            () =>
              `${result.locations[0].logicalLocations[0].fullyQualifiedName} ${tryOr(() => {
                const { index } = result.locations[0].physicalLocation.artifactLocation;
                return result.run.artifacts[index].description.text;
              }, '')}`,
            () => result.locations[0].physicalLocation.artifactLocation.uri,
            '',
          ),
        sortString: (result: Result) =>
          tryOr<string>(
            () => result.locations[0].physicalLocation.artifactLocation.uri,
            '\u2014', // Using escape as VS Packaging munges the char.
          ),
        width: -3,
      },
    ];

    if (this.showAge && this.groupByAge.get()) {
      columns.push({
        id: 'Rule',
        text: '规则',
        filterString: (result: Result) => {
          const rule = result._rule;
          return `${rule.id || rule.guid} ${rule.name ?? ''}`;
        },
        sortString: (result: Result) => {
          const rule = result._rule;
          return `${rule.id || rule.guid} ${rule.name ?? ''}`;
        },
        width: -2,
      });
    }

    columns.push({
      id: 'Details',
      text: '详情',
      filterString: (result: Result) => {
        // TODO: Support templated messages.
        const message = tryOr<string>(
          () => result.message.markdown,
          () => result.message.text, // Can be a constant?
          '',
        );
        const snippet = tryOr<string>(
          () => result.locations[0].physicalLocation.contextRegion.snippet.text,
          () => result.locations[0].physicalLocation.region.snippet.text,
          '',
        );
        return `${message} ${snippet}`;
      },
      sortString: (result: Result) => (result.message.text as string) || '',
      width: -5,
    });

    if (!this.hideBaseline) {
      columns.push({
        id: 'Baseline',
        text: '基线',
        filterString: (result: Result) => (result.baselineState as string) || 'new',
        sortString: (result: Result) => (result.baselineState as string) || 'new',
        width: -1,
      });
    }

    if (!this.hideLevel) {
      columns.push({
        id: 'Level',
        text: '等级',
        filterString: (result: Result) => (result.level as string) || 'error',
        sortString: (result: Result) => (result.level as string) || 'error',
        width: -1,
      });
    }

    const hasWorkItemUris =
      this.run.results &&
      this.run.results.some((result) => result.workItemUris && !!result.workItemUris.length);
    if (hasWorkItemUris) {
      columns.push({
        id: 'Bug',
        text: '错误',
        filterString: (result: Result) => '',
        sortString: (result: Result) => '',
        width: -1,
      });
    }

    if (this.showAge && !this.groupByAge.get()) {
      columns.push({
        id: 'Age',
        text: '年龄',
        filterString: (result: Result) => result.sla,
        sortString: (result: Result) => result.sla,
        width: -1,
      });
    }

    if (this.showAge) {
      columns.push({
        text: 'Observed',
        id: 'First Observed', // Consider using name instead of id
        filterString: (result: Result) => result.firstDetection.toLocaleDateString(),
        sortString: (result: Result) => result.firstDetection.getTime().toString(),
        width: -1,
      });
    }

    return columns;
  }
}
