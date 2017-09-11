import * as React from 'react';
import { Checkbox, DefaultButton  } from 'office-ui-fabric-react';

export interface AddIssueDialogProps {
}

export interface AddIssueDialogState {
    selectedValues: any[][];
}

export class AddIssueDialog extends React.Component<AddIssueDialogProps, AddIssueDialogState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            selectedValues: new Array()
        };
    }

    componentDidMount() {
        var selectedValues: any[][] = JSON.parse(localStorage.getItem('selected-values'));
        console.log('selected-values', selectedValues.length);
        this.setState({ selectedValues });
    }
        
    render() {
        let { selectedValues } = this.state;
        for (var i = 0; i < selectedValues.length; i++) {
        }
        return (
            <div>
                <table>
                    <thead>
                        <tr>
                            <th>種別</th>
                            <th>件名</th>
                            <th>詳細</th>
                            <th>優先度</th>
                        </tr>
                    </thead>
                    <tbody>
                        {selectedValues.map(row => 
                        <tr>
                            <td></td>
                            <td>{row[0]}</td>
                            <td>{row.length > 1 ? row[1]: ''}</td>
                            <td></td>
                        </tr>)}
                    </tbody>
                </table>
            </div>
        );
    };
};
