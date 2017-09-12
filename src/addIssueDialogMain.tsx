import * as React from 'react';
import { render } from 'react-dom';
import { Progress } from './components/progress';
import './assets/styles/global.scss';
import { AddIssueDialog, AddIssueDialogProps } from './components/addIssueDialog';

(() => {
    const title = 'BacklogExcelWebAddin';
    const container = document.querySelector('#container');

    console.log(localStorage.getItem('add-issue-dialog-params'));
    const props: AddIssueDialogProps = JSON.parse(localStorage.getItem('add-issue-dialog-params'));
    
    /* Render application after Office initializes */
    Office.initialize = () => {
        render(
            <AddIssueDialog {...props} />,
            container
        );
    };

    /* Initial render showing a progress bar */
    render(<Progress title={title} logo='assets/logo-filled.png' message='Please sideload your addin to see app body.' />, container);
})();

