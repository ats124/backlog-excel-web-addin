import * as React from 'react';
import * as update from 'immutability-helper'
import { PrimaryButton, DefaultButton, Dropdown, IDropdownOption, DetailsList, IColumn, IGroup, DetailsListLayoutMode, SelectionMode, TextField } from 'office-ui-fabric-react';
import { BacklogProject, BacklogApiKey } from './backlog-project-selector';
import * as backlogjs from 'backlog-js';

export enum ChildParentType {
    Parents,
    Children,
    FirstParentAndChildren
}

export interface AddIssueDialogProps {
    childParentType: ChildParentType;
    selectedApiKey: BacklogApiKey;
    selectedProject: BacklogProject;
    selectedValues: any[][];
}

export interface AddIssueDialogState {
    issues: backlogjs.Option.Issue.PostIssueParams[];
    issueTypeOptions: IDropdownOption[];
    priorityOptions: IDropdownOption[];
    issueDetailListCoumns: IColumn[];
    issueDetailListGroups: IGroup[];
}

export class AddIssueDialog extends React.Component<AddIssueDialogProps, AddIssueDialogState> {
    constructor(props, context) {
        super(props, context);
        let { selectedProject, selectedValues, childParentType } = this.props;

        var issues: backlogjs.Option.Issue.PostIssueParams[] = selectedValues.map(row => ({
            projectId: selectedProject.projectId,
            summary: row[0],
            description: row.length > 1 ? row[1] : '',
            priorityId: selectedProject.priorities[0].id,
            issueTypeId: selectedProject.issueTypes[0].id,
        }));

        let issueTypeOptions: IDropdownOption[] = selectedProject.issueTypes.map(x => ({
            key: x.id,
            text: x.name
        }));

        let priorityOptions: IDropdownOption[] = selectedProject.priorities.map(x => ({
            key: x.id,
            text: x.name
        }));

        let issueDetailListCoumns: IColumn[] = [
            {
                key: 'column1',
                name: '種別',
                fieldName: 'issueTypeId',
                minWidth: 100,
                maxWidth: 140,
                data: 'number',
                onRender: (item, index) => 
                    <Dropdown 
                        options={issueTypeOptions} 
                        selectedKey={item.issueTypeId} 
                        onChanged={(item) => 
                            this.setState({ issues: update(this.state.issues, {[index]: {issueTypeId: {$set: item.key}}}) })
                        }
                    />
            },
            {
                key: 'column2',
                name: '優先度',
                fieldName: 'priorityId',
                minWidth: 100,
                maxWidth: 140,
                data: 'number',
                onRender: (item, index) => 
                    <Dropdown 
                        options={priorityOptions} 
                        selectedKey={item.priorityId} 
                        onChanged={ item => this.setState({ issues: update(this.state.issues, { [index]: { priorityId: { $set: item.key }}})}) }
                    />
            },                    
            {
                key: 'column3',
                name: '件名',
                fieldName: 'summary',
                minWidth: 200,
                maxWidth: 300,
                data: 'string',
                onRender: (item, index) => 
                    <TextField 
                        value={ item.summary } 
                        onChanged={ newValue => this.setState({ issues: update(this.state.issues, { [index]: { summary: { $set: newValue }}})}) } />
            },
            {
                key: 'column4',
                name: '詳細',
                fieldName: 'description',
                minWidth: 300,
                maxWidth: 400,
                data: 'string',
                onRender: (item, index) => 
                <TextField 
                    multiline={ true }
                    value={ item.description }
                    onChanged={ newValue => this.setState({ issues: update(this.state.issues, { [index]: { description: { $set: newValue }}})}) } />
            }
        ];

        // 先頭が親課題でない子課題登録の場合は親課題用の空issueを先頭に挿入する
        if (childParentType == ChildParentType.Children) {
            issues.unshift({
                projectId: selectedProject.projectId,
                summary: '',
                description: '',
                priorityId: selectedProject.priorities[0].id,
                issueTypeId: selectedProject.issueTypes[0].id,
            });
        }

        // 子課題登録する場合はグループを作る
        let issueDetailListGroups: IGroup[] = null;
        if (childParentType == ChildParentType.Children || childParentType == ChildParentType.FirstParentAndChildren && issues.length > 1) {
            issueDetailListGroups = [
                {
                    key: 'parent',
                    name: '親課題',
                    startIndex: 0,
                    count: 1,
                },
                {
                    key: 'children',
                    name: '子課題',
                    startIndex: 1,
                    count: issues.length - 1,
                },
            ];
        }

        this.state = {
            issues, issueTypeOptions, priorityOptions, issueDetailListCoumns, issueDetailListGroups
        };
    }

    componentDidMount() {
    }

    async registButtonOnClick() {
        const { selectedApiKey, childParentType } = this.props;
        const { issues } = this.state;
        const backlog = new backlogjs.Backlog({ host:selectedApiKey.host, apiKey: selectedApiKey.apiKey });

        // 子課題登録する場合は先に親課題を登録してidを取得し
        // 子課題の親課題idをセットする
        if (childParentType != ChildParentType.Parents) {
            const parentIssue = await backlog.postIssue(issues.shift());
            issues.forEach(x => x.parentIssueId = parentIssue.id);
        }

        await Promise.all(issues.map(async x => await backlog.postIssue(x)));
        
        Office.context.ui.messageParent(true);
    }

    cancelButtonOnClick() {
        Office.context.ui.messageParent(false);
    }

    render() {
        let { issues, issueDetailListCoumns, issueDetailListGroups } = this.state;
        return (
            <div>
                <DetailsList
                    items={ issues }
                    columns={ issueDetailListCoumns }
                    layoutMode={ DetailsListLayoutMode.justified }
                    isHeaderVisible={ true }
                    selectionMode={ SelectionMode.none }
                    groups={ issueDetailListGroups }
                />
                <footer className='ms-u-textAlignRight'>
                    <PrimaryButton onClick={this.registButtonOnClick.bind(this)}>登録</PrimaryButton>
                    <DefaultButton onClick={this.cancelButtonOnClick.bind(this)}>キャンセル</DefaultButton>
                </footer>
            </div>
        );
    };
};
