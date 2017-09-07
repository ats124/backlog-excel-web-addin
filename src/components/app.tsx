import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import { Header } from './header';
import { HeroList, HeroListItem } from './hero-list';
import { BacklogProjectSelector } from './backlog-project-selector';

export interface AppProps {
    title: string;
}

export interface AppState {
    listItems: HeroListItem[];
}

export class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            listItems: []
        };
    }

    componentDidMount() {
        this.setState({
            listItems: [
                {
                    icon: 'Ribbon',
                    primaryText: 'Achieve more with Office integration'
                },
                {
                    icon: 'Unlock',
                    primaryText: 'Unlock features and functionality'
                },
                {
                    icon: 'Design',
                    primaryText: 'Create and visualize like a pro'
                }
            ]
        });
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
            <div className='ms-welcome'>
                <BacklogProjectSelector />
            </div>
        );
    };
};
