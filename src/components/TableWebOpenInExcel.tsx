// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React, { ReactNode, useState } from 'react';
import { PrimaryButton, Stack } from '@fluentui/react';
import ReactDOMServer from 'react-dom/server';
import workbookManager from '..';

interface TableWebOpenInExcelProps {
    children: ReactNode;
}

const TableWebOpenInExcel: React.FC<TableWebOpenInExcelProps> = ({ children }) => {
    const [isHovered, setIsHovered] = useState(false);
    return (
        <div
            className="wrapper"
            onMouseEnter={() => setIsHovered(true)}
            onMouseLeave={() => setIsHovered(false)}
        >
            <Stack horizontalAlign="center">
                {isHovered && (
                    <PrimaryButton
                        className="hoverButton"
                        text="Open in Excel"
                        onClick={() =>
                            exportTableFromHtml(
                                'Table' + '.xlsx',
                                ReactDOMServer.renderToStaticMarkup(children as React.ReactElement)
                            )
                        }
                        allowDisabledFocus
                    />
                )}
                {children}
            </Stack>
        </div>
    );
};

const exportTableFromHtml = async (filename: string, HTMLTableString: string) => {
    const parser = new DOMParser();
    const htmlDoc = parser.parseFromString(HTMLTableString, 'text/html');
    const blob = await workbookManager.generateTableWorkbookFromHtml(
        htmlDoc.querySelector('table') as HTMLTableElement
    );
    workbookManager.openInExcelWeb(blob, filename);
};

export default TableWebOpenInExcel;
