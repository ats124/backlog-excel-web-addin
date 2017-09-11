import * as React from 'react';
import { Checkbox, DefaultButton  } from 'office-ui-fabric-react';
import { BacklogProjectSelector, BacklogApiKey, BacklogProject } from './backlog-project-selector';

export interface AppProps {
    title: string;
}

export interface AppState {
    selectedApiKey: BacklogApiKey;
    selectedProject: BacklogProject;
    isChildren: boolean;

}

export class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            selectedApiKey: null,
            selectedProject: null,
            isChildren: false,
        };
    }

    componentDidMount() {
    }

    addIssuesButtonOnClick = async() => {
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

            console.log('selectedValues', selectedValues);
            localStorage.setItem('selected-values', JSON.stringify(selectedValues));
            Office.context.ui.displayDialogAsync('https://localhost:3000/addIssueDialog.html', { width: 50, height: 50, xFrameDenySafe: true });  
        });
    }
        
    render() {
        let { selectedProject, isChildren } = this.state;
        return (
            <div>
                <BacklogProjectSelector onChanged={((apiKey, project) => this.setState({ selectedApiKey: apiKey, selectedProject: project })) } />
                <Checkbox 
                    label='子課題として登録する'
                    checked={isChildren} 
                    onChange={ ((_, isChecked) => this.setState({ isChildren: isChecked })) } />
                <DefaultButton onClick={this.addIssuesButtonOnClick} disabled={selectedProject == null}>選択範囲を課題として登録</DefaultButton>
            </div>
        );
    };
};
