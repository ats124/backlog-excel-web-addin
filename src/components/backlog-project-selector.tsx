import * as React from 'react';
import * as backlogjs from 'backlog-js';
import { Dropdown, IDropdownOption, Dialog, DialogFooter, TextField, PrimaryButton, DefaultButton } from 'office-ui-fabric-react';

export interface BacklogProject {
    projectId: number;
    name: string;
    priorities: BacklogPriority[];
    issueTypes: BacklogIssueType[];
}

export interface BacklogPriority {
    id: number;
    name: string;
}

export interface BacklogIssueType {
    id: number;
    name: string;
    displayOrder: number;
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
    firstLoadProjects: boolean,
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
            firstLoadProjects: true,
        }
    }

    componentDidMount() {
        // 設定からAPIキーを取得する
        var apiKeys: BacklogApiKey[] = Office.context.document.settings.get('backlog-api-keys');
        if (apiKeys == null) apiKeys = [];

        // 設定から最後から選択したAPIキーを取得する
        var selectedApiKey: BacklogApiKey;
        var lastSelectedApiKeyName: string = Office.context.document.settings.get('backlog-last-selected-apikey');
        if (lastSelectedApiKeyName != null && lastSelectedApiKeyName != undefined)
        {
            selectedApiKey = apiKeys.find(x => x.name == lastSelectedApiKeyName);
            if (selectedApiKey == undefined) selectedApiKey = null;
        } else {
            selectedApiKey = null;
        }

        this.setState({ apiKeys: apiKeys, selectedApiKey });
        this.props.onChanged(selectedApiKey, null);
        
        // 最後の選択したAPIキーがある場合はプロジェクトを読み込む
        if (selectedApiKey != null) {
            this.reloadProjects(selectedApiKey);
        }
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

        // APIキー一覧を更新
        var newApiKey: BacklogApiKey = { name: inputApiKeyName, host: inputHost, apiKey: inputApiKey };
        apiKeys.push(newApiKey);
        Office.context.document.settings.set('backlog-api-keys', apiKeys);
        await Office.context.document.settings.saveAsync();
        
        // 追加したAPIキーを選択
        this.setState({ apiKeys, selectedApiKey: newApiKey });
        this.closeDialog();

        // プロジェクト再読込
        this.reloadProjects(newApiKey);
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
        
        this.setState({ apiKeys, selectedApiKey: null, projects: [], selectedProject: null });
        this.props.onChanged(null, null);
    }

    apiKeyDropdownOnChanged = async(_, index) => {
        var selectedApiKey = this.state.apiKeys[index];
        this.setState({ selectedApiKey, selectedProject:null, projects: [] });
        this.props.onChanged(selectedApiKey, null);

        Office.context.document.settings.set('backlog-last-selected-apikey', selectedApiKey.name);
        await Office.context.document.settings.saveAsync();
        
        await this.reloadProjects(selectedApiKey);
    }

    projectDropdownOnChanged = async(_, index) => {
        let { selectedApiKey, projects } = this.state;
        var selectedProject = projects[index];
        this.setState({ selectedProject });
        this.props.onChanged(selectedApiKey, selectedProject);

        Office.context.document.settings.set('backlog-last-selected-project', selectedProject.projectId);
        await Office.context.document.settings.saveAsync();
    }
 
    closeDialog() {
        this.setState({ isApiKeyDialogOpen: false });
    }

    async reloadProjects(selectedApiKey: BacklogApiKey) {        
        if (selectedApiKey != null) {
            const backlog = new backlogjs.Backlog({ host: selectedApiKey.host, apiKey: selectedApiKey.apiKey });
            try {
                const priorities: BacklogPriority[] = (await backlog.getPriorities()).map(x => ({id: x.id, name: x.name}));
                const projectsData = await backlog.getProjects();
                var projects: BacklogProject[] = [];
                for (var i = 0; i < projectsData.length; i++) {
                    const issueTypes: BacklogIssueType[] = (await backlog.getIssueTypes(projectsData[i].id)).map(x => ({id: x.id, name: x.name, displayOrder:x.displayOrder}));
                    projects.push({ projectId: projectsData[i].id, name: projectsData[i].name, priorities, issueTypes });
                }

                // 初回のプロジェクト読み込み時は設定から最後に選択したプロジェクトを選択する
                var selectedProject: BacklogProject = null;
                if (this.state.firstLoadProjects) {
                    var lastSelectedProjectId = Office.context.document.settings.get('backlog-last-selected-project');
                    if (lastSelectedProjectId != undefined) {
                        selectedProject = projects.find(x => x.projectId == lastSelectedProjectId);
                        if (selectedProject == undefined) selectedProject = null;
                    }
                    this.setState({firstLoadProjects: false});
                }

                // 最後に選択したプロジェクトを選択できなければ先頭のプロジェクトを選択
                if (selectedProject == null && projects.length > 0) {
                    selectedProject = projects[0];
                }
                
                this.setState({ projects, selectedProject });
                this.props.onChanged(selectedApiKey, selectedProject);

                Office.context.document.settings.set('backlog-last-selected-project', selectedProject.projectId);
                await Office.context.document.settings.saveAsync();         
                
            } catch (e) {
                this.clearProjects();                
            }
        } else {
            this.clearProjects();
        }
    }

    clearProjects() {
        this.setState({ projects: [] });
        if (this.state.selectedProject != null) {
            this.setState({selectedProject: null});
            this.props.onChanged(null, null);
        }
    }

    render() {
        let { apiKeys, selectedApiKey, selectedProject, inputApiKeyName, inputHost, inputApiKey } = this.state;

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
                    onChanged={this.projectDropdownOnChanged}
                    isDisabled={projectItems.length == 0} />
            </div>
        );
    };
};
