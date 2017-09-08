import * as React from 'react';
import * as backlogjs from 'backlog-js';
import { Dropdown, IDropdownOption, Dialog, DialogFooter, TextField, PrimaryButton, DefaultButton } from 'office-ui-fabric-react';

export interface BacklogProject {
    projectId: number;
    name: string;
}

export interface BacklogApiKey {
    name: string;
    host: string;
    apiKey: string;
}

export interface BacklogProjectSelectorProps {
    onChanged?: (apiKey?: BacklogApiKey, project?: BacklogProject) => void;
}

export interface BacklogProjectSelectorState {
    apiKeys: BacklogApiKey[];
    projects: BacklogProject[];
    selectedApiKey: BacklogApiKey;
    selectedProject: BacklogProject;
    isApiKeyDialogOpen: boolean;
    inputApiKeyName: string;
    inputHost: string;
    inputApiKey: string;
}

export class BacklogProjectSelector extends React.Component<BacklogProjectSelectorProps, BacklogProjectSelectorState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            apiKeys: [ ],
            projects: [ ],
            selectedApiKey: null,
            selectedProject: null,
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
        apiKeys.push({ name: inputApiKeyName, host: inputHost, apiKey: inputApiKey });
        Office.context.document.settings.set('backlog-api-keys', apiKeys);
        await Office.context.document.settings.saveAsync();
        
        this.setState({ apiKeys });
        this.closeDialog();
    }

    deleteApiKeyButtonClick = async () => {
        let { apiKeys, selectedApiKey } = this.state;
        var newSelectedApiKey: BacklogApiKey = null;
        if (apiKeys.length > 1) {
            var i = apiKeys.indexOf(selectedApiKey);
            if (i > 0) newSelectedApiKey = apiKeys[i - 1];
            else newSelectedApiKey = apiKeys[0];
        }

        apiKeys = apiKeys.filter(x => x != selectedApiKey);
        Office.context.document.settings.set('backlog-api-keys', apiKeys);
        await Office.context.document.settings.saveAsync();
        
        this.setState({ apiKeys, selectedApiKey: null, projects: [], selectedProject: null });
        this.props.onChanged(null, null);
    }
 
    closeDialog() {
        this.setState({ isApiKeyDialogOpen: false });
    }

    apiKeyDropdownOnChanged = async(_, index) => {
        this.setState({ selectedProject:null,  projects: [] });
        if (index >= 0) {
            this.setState({ selectedApiKey: this.state.apiKeys[index] });
            this.props.onChanged(this.state.apiKeys[index], null);
            let { host, apiKey } = this.state.apiKeys[index];
            const backlog = new backlogjs.Backlog({ host, apiKey });
            backlog.getProjects().then(data => {
                var projects: BacklogProject[] = [];
                for (var i = 0; i < data.length; i++) {
                    projects.push({ projectId: data[i].id, name: data[i].name });
                }
                this.setState({ projects });
            }).catch(err => {
                this.setState({ projects: [] });
            });
        }
        else
        {
            this.setState({ selectedApiKey: null});
            this.props.onChanged(null, null);
        }
    }

    projectDropdownOnChanged = async(_, index) => {
        if (index >= 0) {
            var apiKey = this.state.apiKeys[index];
            const backlog = new backlogjs.Backlog({host:apiKey.host, apiKey: apiKey.apiKey});
            backlog.getProjects().then(data => {
                var projects: BacklogProject[] = [];
                for (var i = 0; i < data.length; i++) {
                    projects.push({ projectId: data[i].id, name: data[i].name });
                }
                this.setState({ projects });
            }).catch(err => {
                this.setState({ projects: [] });
            });
        }
        else
        {
            this.setState({ selectedProject: null});
            this.props.onChanged(null, null);
        }
    }

    render() {
        let {apiKeys, selectedApiKey, selectedProject, inputApiKeyName, inputHost, inputApiKey} = this.state;

        var keyItems: IDropdownOption[] = [];
        var selectedApiOptionKey: number = null;
        for (var i = 0; i < apiKeys.length; i++) {
            var apiKey = apiKeys[i];
            keyItems.push({key: i, text: apiKey.name});
            if (apiKey == selectedApiKey) selectedApiOptionKey = i;
        }
        var projectItems: IDropdownOption[] = [];
        var selectedProjectOptionKey: number = null;
        for (var i = 0; i < this.state.projects.length; i++) {
            var p = this.state.projects[i];
            projectItems.push({key: i, text: p.name});
            if (p == selectedProject) selectedProjectOptionKey = i;
        }
        
        return (
            <div>
                <Dropdown
                    label='APIキー'
                    placeHolder='APIキーを選択してください'
                    options={keyItems} 
                    onChanged={this.apiKeyDropdownOnChanged}
                    selectedKey={selectedApiOptionKey}/>
                <div className="ms-u-textAlignRight">
                    <DefaultButton onClick={this.addApiKeyButtonClick}>追加</DefaultButton>
                    <DefaultButton onClick={this.deleteApiKeyButtonClick} disabled={selectedApiKey == null}>削除</DefaultButton>
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
                    selectedKey={selectedProjectOptionKey}
                    isDisabled={projectItems.length == 0} />
            </div>
        );
    };
};
