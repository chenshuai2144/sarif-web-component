// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import './RunCard.scss';
import * as React from 'react';
import { Component } from 'react';
import { autorun, runInAction, observable, computed, untracked } from 'mobx';
import { observer } from 'mobx-react';

import { Hi } from './Hi';
import { renderCell } from './RunCard.renderCell';
import { More, ResultOrRuleOrMore } from './Viewer.Types';
import { RunStore, SortRuleBy } from './RunStore';
import { TreeColumnSorting } from './RunCard.TreeColumnSorting';
import { tryOr } from './try';

import { Card } from 'azure-devops-ui/Card';
import { Observer } from 'azure-devops-ui/Observer';
import { ObservableValue, IObservableValue } from 'azure-devops-ui/Core/Observable';
import { IHeaderCommandBarItem } from 'azure-devops-ui/HeaderCommandBar';
import { MenuItemType } from 'azure-devops-ui/Menu';
import { Pill, PillSize } from 'azure-devops-ui/Pill';
import { Tree, ITreeColumn } from 'azure-devops-ui/TreeEx';
import { TreeItemProvider, ITreeItemEx } from 'azure-devops-ui/Utilities/TreeItemProvider';
import { Tooltip } from 'azure-devops-ui/TooltipEx';

@observer
export class RunCard extends Component<{ runStore: RunStore; index: number; runCount: number }> {
  @observable private show = true;
  private groupByMenuItems = [] as IHeaderCommandBarItem[];
  private itemProvider = new TreeItemProvider<ResultOrRuleOrMore>([]);
  private columnCache = new Map<string, ITreeColumn<ResultOrRuleOrMore>>();

  @computed({ keepAlive: true }) private get sortRuleByMenuItems(): IHeaderCommandBarItem[] {
    const { runStore } = this.props;
    const sortRuleBy = untracked(() => runStore.sortRuleBy);
    const onActivate = (menuItem) => {
      runStore.sortRuleBy = menuItem.data;
      this.sortRuleByMenuItems.forEach(
        (item) => ((item.checked as IObservableValue<boolean>).value = item.id === menuItem.id),
      );
    };
    return [
      {
        data: SortRuleBy.Count,
        id: 'sortByRuleCount',
        text: '按照规则数量排序',
        ariaLabel: '按照规则数量排序',
        onActivate,
        important: false,
        checked: new ObservableValue(sortRuleBy === SortRuleBy.Count),
      },
      {
        data: SortRuleBy.Name,
        id: 'sortByRuleName',
        text: '按照规则名称排序',
        ariaLabel: '按照规则名称排序',
        onActivate,
        important: false,
        checked: new ObservableValue(sortRuleBy === SortRuleBy.Name),
      },
    ];
  }

  @computed private get columns() {
    const { runStore } = this.props;
    return runStore.columns.map((col, i) => {
      const { id, width, text } = col;
      if (!this.columnCache.has(id)) {
        const observableWidth = new ObservableValue(width);
        this.columnCache.set(id, {
          id: id.replace(/ /g, ''),
          name: text,
          width: observableWidth,
          onSize: (e, i, newWidth) => (observableWidth.value = newWidth),
          renderCell: renderCell,
          sortProps: {
            ariaLabelAscending: '按 A 到 Z 排序',
            ariaLabelDescending: '从 Z 到 A 排序',
            sortOrder: i === runStore.sortColumnIndex ? runStore.sortOrder : undefined,
          },
        } as unknown as ITreeColumn<ResultOrRuleOrMore>);
      }
      return this.columnCache.get(id);
    });
  }

  constructor(props) {
    super(props);
    const { runStore } = this.props;

    if (runStore.showAge) {
      const onActivateGroupBy = (menuItem) => {
        runStore.groupByAge.set(menuItem.data);
        this.groupByMenuItems
          .filter((item) => item.itemType !== MenuItemType.Divider)
          .forEach(
            (item) => ((item.checked as IObservableValue<boolean>).value = item.id === menuItem.id),
          );
      };

      this.groupByMenuItems = [
        {
          data: true,
          id: 'groupByAge',
          text: 'Group by age',
          ariaLabel: 'Group by age',
          onActivate: onActivateGroupBy,
          important: false,
          checked: new ObservableValue(runStore.groupByAge.get()),
        },
        {
          data: false,
          id: 'groupByRule',
          text: '按规则分组',
          ariaLabel: '按规则分组',
          onActivate: onActivateGroupBy,
          important: false,
          checked: new ObservableValue(!runStore.groupByAge.get()),
        },
        { id: 'separator', important: false, itemType: MenuItemType.Divider },
      ];
    }

    autorun(() => {
      this.itemProvider.clear();
      this.itemProvider.splice(undefined, undefined, [
        { items: this.props.runStore.rulesTruncated },
      ]);
    });

    autorun(() => (this.show = this.props.index === 0));
  }

  private sortingBehavior = new TreeColumnSorting<ITreeItemEx<ResultOrRuleOrMore>>(
    (columnIndex, proposedSortOrder) => {
      for (let index = 0; index < this.columns.length; index++) {
        const column = this.columns[index];
        if (column.sortProps) {
          column.sortProps.sortOrder = index === columnIndex ? proposedSortOrder : undefined;
        }
      }
      runInAction(() => {
        this.props.runStore.sortColumnIndex = columnIndex;
        this.props.runStore.sortOrder = proposedSortOrder;
      });
    },
  );

  render() {
    const { show, itemProvider } = this;
    const { runStore, runCount } = this.props;

    return (
      <Observer renderChildren={itemProvider}>
        {() => {
          const qualityDomain = tryOr(
            () => runStore.run.tool.driver.properties['microsoft/qualityDomain'],
          );
          return (
            <Card
              titleProps={{
                ariaLevel: 2,
                text: (
                  <Tooltip
                    text={
                      (
                        <>
                          <div>
                            {tryOr(
                              () => runStore.run.tool.driver.fullName,
                              () =>
                                `${runStore.run.tool.driver.name} ${
                                  runStore.run.tool.driver.semanticVersion || ''
                                }`,
                            )}
                          </div>
                          {tryOr(
                            () => (
                              <div>{runStore.run.tool.driver.fullDescription.text}</div>
                            ),
                            () => (
                              <div>{runStore.run.tool.driver.shortDescription.text}</div>
                            ),
                          )}
                        </>
                      ) as any
                    }
                  >
                    <span className={'swcRunTitle'}>
                      <Hi>{runStore.driverName}</Hi>
                      {qualityDomain && ` (${qualityDomain})`}
                      <Pill size={PillSize.compact}>{runStore.filteredCount}</Pill>
                    </span>
                    {/* Tooltip marked as React.Children.only thus extra span. */}
                  </Tooltip>
                ) as any,
              }}
              contentProps={{ contentPadding: false }}
              headerCommandBarItems={[
                runCount > 1
                  ? {
                      id: 'hide',
                      text: '', // Remove?
                      ariaLabel: '显示/隐藏',
                      onActivate: () => (this.show = !this.show),
                      iconProps: { iconName: this.show ? 'ChevronDown' : 'ChevronUp' }, // Naturally updates as this entire object is re-created each render.
                      important: runCount > 1,
                    }
                  : undefined,
                ...this.groupByMenuItems,
                ...this.sortRuleByMenuItems,
              ].filter((item) => item)}
              className="flex-grow bolt-card-no-vertical-padding"
            >
              {show &&
                (itemProvider.length ? (
                  <Tree<ResultOrRuleOrMore>
                    className="swcTree"
                    columns={this.columns}
                    itemProvider={itemProvider}
                    onToggle={(event, treeItem: ITreeItemEx<ResultOrRuleOrMore>) => {
                      itemProvider.toggle(treeItem.underlyingItem);
                    }}
                    onActivate={(event, treeRow) => {
                      const treeItem = treeRow.data.underlyingItem;
                      const more = treeItem.data as More;
                      if (more.onClick) {
                        more.onClick(); // Handle "Show All"
                      } else {
                        itemProvider.toggle(treeItem);
                      }
                    }}
                    behaviors={[this.sortingBehavior]}
                    selectableText={true}
                  />
                ) : (
                  <div className="swcRunEmpty">No Results</div>
                ))}
            </Card>
          );
        }}
      </Observer>
    );
  }
}
