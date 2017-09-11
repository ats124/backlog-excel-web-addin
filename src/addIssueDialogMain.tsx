import * as React from 'react';
import { render } from 'react-dom';
import { Progress } from './components/progress';
import './assets/styles/global.scss';
import { AddIssueDialog } from './components/addIssueDialog';

(() => {
    const title = 'BacklogExcelWebAddin';
    const container = document.querySelector('#container');

    /* Render application after Office initializes */
    Office.initialize = () => {
        render(
            <AddIssueDialog />,
            container
        );
    };

    /* Initial render showing a progress bar */
    render(<Progress title={title} logo='assets/logo-filled.png' message='Please sideload your addin to see app body.' />, container);
})();

