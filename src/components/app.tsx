import * as React from 'react';
import { Checkbox, DefaultButton  } from 'office-ui-fabric-react';
import { BacklogProjectSelector, BacklogApiKey, BacklogProject } from './backlog-project-selector';
import * as backlogjs from 'backlog-js';

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
            isChildren: false
        };
    }

    componentDidMount() {
    }

    addIssuesButtonOnClick = async() => {
        let { selectedApiKey, selectedProject, isChildren } = this.state;

        Excel.run(function (ctx) {
            // 選択された範囲に対するプロキシ オブジェクトを作成し、そのプロパティを読み込みます
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");
            
            // キューに入れるコマンドを実行し、タスクの完了を示すために Promise を返します
            return ctx.sync()
                .then(function () {
                    const backlog = new backlogjs.Backlog({host: selectedApiKey.host, apiKey: selectedApiKey.apiKey});
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        var params: backlogjs.Option.Issue.PostIssueParams = { projectId: selectedProject.projectId, summary: sourceRange.values[i][0], priorityId: 1 };
                        if (sourceRange.columnCount > 1) {
                            params.description = sourceRange.values[i][1];
                        }
                        backlog.postIssue(params);
                    }
                    backlogjs.Option.Issue
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // セルを検索して強調表示します
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        for (var j = 0; j < sourceRange.columnCount; j++) {
                            if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                highestRow = i;
                                highestCol = j;
                                highestValue = sourceRange.values[i][j];
                            }
                        }
                    }

                    cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                    sourceRange.worksheet.getUsedRange().format.fill.clear();
                    sourceRange.worksheet.getUsedRange().format.font.bold = false;

                    // セルを強調表示
                    cellToHighlight.format.fill.color = "orange";
                    cellToHighlight.format.font.bold = true;
                })
                .then(ctx.sync);
    })
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
