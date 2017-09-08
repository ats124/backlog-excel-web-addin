import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import { Header } from './header';
import { BacklogProjectSelector, BacklogApiKey, BacklogProject } from './backlog-project-selector';

export interface AppProps {
    title: string;
}

export interface AppState {
    selectedApiKey: BacklogApiKey;
    selectedProject: BacklogProject;
}

export class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            selectedApiKey: null,
            selectedProject: null
        };
    }

    componentDidMount() {
    }

    click = async () => {
        
        await Excel.run(async (context) => {
            /**
             * Insert your Excel code here
             */
            await context.sync();
        });
        
    }

    render() {
        return (
            <BacklogProjectSelector onChanged={((apiKey, project) => this.setState({ selectedApiKey: apiKey, selectedProject: project })) } />
        );
    };
};
