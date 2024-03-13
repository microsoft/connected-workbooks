// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React, { ReactNode, useState } from 'react';
import { PrimaryButton, Stack } from '@fluentui/react';
import { exportTableFromHtml } from '../workbook-api/api';
import ReactDOMServer from 'react-dom/server';

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

export default TableWebOpenInExcel;
