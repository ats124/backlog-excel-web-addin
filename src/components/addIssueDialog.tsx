import * as React from 'react';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react';
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

        this.state = {
            issues
        };
    }

    componentDidMount() {
    }
        
    render() {
        let { issues } = this.state;
        return (
            <div>
                <table className='ms-Table'>
                    <thead>
                        <tr>
                            <th>種別</th>
                            <th>件名</th>
                            <th>詳細</th>
                            <th>優先度</th>
                        </tr>
                    </thead>
                    <tbody>
                        {issues.map(issue => 
                        <tr>
                            <td></td>
                            <td>{issue.summary}</td>
                            <td>{issue.description}</td>
                            <td></td>
                        </tr>)}
                    </tbody>
                </table>
                <footer className='ms-u-textAlignRight'>
                    <PrimaryButton>登録</PrimaryButton>
                    <DefaultButton>キャンセル</DefaultButton>
                </footer>
            </div>
        );
    };
};
