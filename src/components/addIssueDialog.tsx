import * as React from 'react';
import * as update from 'immutability-helper'
import { PrimaryButton, DefaultButton, Dropdown, IDropdownOption, DetailsList, IColumn, DetailsListLayoutMode, SelectionMode, TextField } from 'office-ui-fabric-react';
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
}

export class AddIssueDialog extends React.Component<AddIssueDialogProps, AddIssueDialogState> {
    constructor(props, context) {
        super(props, context);
        let {selectedProject, selectedValues} = this.props;

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

        this.state = {
            issues, issueTypeOptions, priorityOptions, issueDetailListCoumns
        };
    }

    componentDidMount() {
    }
        
    render() {
        let { issues, issueDetailListCoumns } = this.state;
        return (
            <div>
                <DetailsList
                    items={ issues }
                    columns={ issueDetailListCoumns }
                    layoutMode={ DetailsListLayoutMode.justified }
                    isHeaderVisible={ true }
                    selectionMode={ SelectionMode.none }
                />
                <footer className='ms-u-textAlignRight'>
                    <PrimaryButton>登録</PrimaryButton>
                    <DefaultButton>キャンセル</DefaultButton>
                </footer>
            </div>
        );
    };
};
