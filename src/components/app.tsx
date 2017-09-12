import * as React from 'react';
import { Label, DefaultButton } from 'office-ui-fabric-react';
import { BacklogProjectSelector, BacklogApiKey, BacklogProject } from './backlog-project-selector';
import { ChildParentType, AddIssueDialogProps } from './addIssueDialog';

declare var __BASE_URL__: string;

export interface AppProps {
    title: string;
}

export interface AppState {
    selectedApiKey: BacklogApiKey;
    selectedProject: BacklogProject;
    childParentType: ChildParentType;
}

export class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            selectedApiKey: null,
            selectedProject: null,
            childParentType: ChildParentType.Parents,
        };
    }

    componentDidMount() {
    }

    addIssuesButtonOnClick = async() => {
        let { selectedApiKey, selectedProject, childParentType } = this.state;

        await Excel.run(async ctx => {
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");            
            await ctx.sync();

            var selectedValues: any[][] = new Array();
            for (var i = 0; i < sourceRange.rowCount; i++) {
                selectedValues[i] = new Array();
                for (var j = 0; j < sourceRange.columnCount; j++) {
                    selectedValues[i][j] = sourceRange.values[i][j];
                }
            }

            const props: AddIssueDialogProps = { selectedApiKey, selectedProject, childParentType, selectedValues };
            localStorage.setItem('add-issue-dialog-params', JSON.stringify(props));
            Office.context.ui.displayDialogAsync(
                __BASE_URL__ + '/addIssueDialog.html', 
                { width: 50, height: 50, xFrameDenySafe: true }, 
                asyncResult => { 
                    const dialog = asyncResult.value;
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, dialogResult => {
                        dialog.close();
                        if (dialogResult) {
                            
                        }
                    });
                });  
        });
    }

    render() {
        let { selectedProject, childParentType } = this.state;
        const childParentTypeRadios = [
            { type: ChildParentType.Parents, text: '親課題として登録する' },
            { type: ChildParentType.Children, text: '子課題として登録する' },
            { type: ChildParentType.FirstParentAndChildren, text: '子課題として登録する(先頭を親課題とする)' },
        ].map(x => <div><label><input type="radio" value={x.type} checked={childParentType == x.type} onChange={() => this.setState({ childParentType: x.type })} />{x.text}</label></div>);
        
        return (
            <div>
                <BacklogProjectSelector onChanged={((apiKey, project) => this.setState({ selectedApiKey: apiKey, selectedProject: project })) } />
                <div>
                    <Label>親子関係</Label>
                    {childParentTypeRadios}
                </div>
                <DefaultButton onClick={this.addIssuesButtonOnClick} disabled={selectedProject == null}>選択範囲を課題として登録</DefaultButton>
            </div>
        );
    };
};
