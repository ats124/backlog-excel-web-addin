import * as React from 'react';
import * as backlogjs from 'backlog-js';
import { Dropdown, IDropdownOption, Dialog, DialogFooter, TextField, PrimaryButton, DefaultButton } from 'office-ui-fabric-react';

export interface BacklogProject {
    projectId: number;
    name: string;
    selected: boolean;
}

export interface BacklogApiKey {
    name: string;
    host: string;
    apiKey: string;
    selected: boolean;
}

export interface BacklogProjectSelectorState {
    apiKeys: BacklogApiKey[];
    projects: BacklogProject[];
    isApiKeyDialogOpen: boolean;
    inputApiKeyName: string;
    inputHost: string;
    inputApiKey: string;
}

export class BacklogProjectSelector extends React.Component<any, BacklogProjectSelectorState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            apiKeys: [ ],
            projects: [ ],
            isApiKeyDialogOpen: false,
            inputApiKeyName: '',
            inputHost: '',
            inputApiKey: '',
        }
    }

    componentDidMount() {
        this.reloadApiKeys();
    }

    reloadApiKeys() {
        var apiKeys: BacklogApiKey[] = Office.context.document.settings.get('backlog-api-keys');
        if (apiKeys == null) apiKeys = [];
        this.setState({ apiKeys: apiKeys });
    }

    addApiKeyButtonClick = async () => {
        this.setState({ 
            inputApiKeyName: '',
            inputHost: '',
            inputApiKey: '',
            isApiKeyDialogOpen: true 
        });
    }

    okApiKeyButtonClick = async () => {
        let { apiKeys, inputApiKeyName, inputHost, inputApiKey } = this.state;
        apiKeys.push({ name: inputApiKeyName, host: inputHost, apiKey: inputApiKey, selected: true });
        Office.context.document.settings.set('backlog-api-keys', apiKeys);
        await Office.context.document.settings.saveAsync();
        
        this.setState({ apiKeys });
        this.closeDialog();
    }

    closeDialog() {
        this.setState({ isApiKeyDialogOpen: false });
    }

    apiKeyDropdownOnChanged = async(_, index) => {
        if (index >= 0) {
            let { host, apiKey } = this.state.apiKeys[index];
            const backlog = new backlogjs.Backlog({ host, apiKey });
            var projects: BacklogProject[] = [];
            backlog.getProjects().then(data => {
                for (var i = 0; i < data.length; i++) {
                    projects.push({ projectId: data.id, name: data.name, selected: false });
                }
                this.setState({ projects })
            }).catch(err => {
            });
        } else {
            this.setState({ projects: [] })
        }
    }

    render() {
        let {apiKeys, inputApiKeyName, inputHost, inputApiKey} = this.state;

        var keyItems: IDropdownOption[] = [];
        for (var i = 0; i < apiKeys.length; i++) {
            var apiKey = apiKeys[i];
            keyItems.push({key: i, text: apiKey.name});
        }
        var projectItems: IDropdownOption[] = [];
        for (var i = 0; i < this.state.projects.length; i++) {
            var p = this.state.projects[0];
            projectItems.push({key: i, text: p.name});
        }
        
        return (
            <div>
                <Dropdown
                    label='APIキー'
                    placeHolder='APIキーを選択してください'
                    options={keyItems} />
                <div>
                    <DefaultButton onClick={this.addApiKeyButtonClick}>追加</DefaultButton>
                    <DefaultButton>削除</DefaultButton>
                </div>
                <Dialog
                    title='APIキーの追加'
                    isOpen={this.state.isApiKeyDialogOpen}
                    onDismiss={this.closeDialog.bind(this)}>
                    <TextField 
                        label='名称'
                        value={inputApiKeyName}
                        onChanged={(text) => this.setState({inputApiKeyName: text}) }/>
                    <TextField 
                        label='ホスト'
                        addonString='https://'
                        value={inputHost}
                        onChanged={(text) => this.setState({inputHost: text}) }/>
                    <TextField
                        label='APIキー'
                        value={inputApiKey}
                        onChanged={(text) => this.setState({inputApiKey: text}) }/>
                    <DialogFooter>
                        <PrimaryButton onClick={this.okApiKeyButtonClick}>OK</PrimaryButton>
                        <DefaultButton onClick={this.closeDialog.bind(this)}>キャンセル</DefaultButton>
                    </DialogFooter>
                </Dialog>
                <Dropdown
                    label='プロジェクト'
                    placeHolder='プロジェクトを選択してください'
                    options={projectItems} 
                    isDisabled={projectItems.length == 0} />
            </div>
        );
    };
};
