define([], function () {
    'use strict';

    return {
        // Properties panel definition
        definition: {
            type: "items",
            component: "accordion",
            items: {
                settings: {
                    uses: "settings",
                    items: {
                        general: {
                            type: "items",
                            label: "General",
                            items: {
                                exampleText: {
                                    ref: "props.exampleText",
                                    label: "Example Text",
                                    type: "string",
                                    expression: "optional",
                                    defaultValue: "Hello from Template_Qlik!"
                                },
                                showBorder: {
                                    ref: "props.showBorder",
                                    label: "Show Border",
                                    type: "boolean",
                                    defaultValue: false
                                }
                            }
                        }
                    }
                },
                data: {
                    uses: "data",
                    items: {
                        dimensions: {
                            uses: "dimensions",
                            min: 0,
                            max: 1
                        },
                        measures: {
                            uses: "measures",
                            min: 0,
                            max: 1
                        }
                    }
                },
                mcpOrchestratorSettings: {
                    type: "items",
                    label: "GSE - MCP Orchestrator",
                    items: {
                        mcpHeader: {
                            component: "text",
                            label: "Configure MCP Orchestrator URL for integration with SharePoint and Claude AI."
                        },
                        mcpOrchestratorUrl: {
                            type: "string",
                            ref: "mcpOrchestratorUrl",
                            label: "MCP Orchestrator URL (no trailing slash)",
                            defaultValue: "https://gse-mcp.replit.app",
                            expression: "optional"
                        },
                        ssoInfo: {
                            component: "text",
                            label: "🔒 Authentication: SSO (session-based via Qlik Cloud)"
                        }
                    }
                },
                mcpClaudeSettings: {
                    type: "items",
                    label: "GSE - MCP Claude",
                    items: {
                        claudeHeader: {
                            component: "text",
                            label: "Configure Claude AI for intelligent data analysis."
                        },
                        claudeSystemPrompt: {
                            type: "string",
                            component: "textarea",
                            ref: "claudeSystemPrompt",
                            label: "System Prompt (Claude's role)",
                            defaultValue: "You are a business intelligence analyst. Answer directly with only the requested information.",
                            rows: 3
                        },
                        claudeMaxTokens: {
                            type: "number",
                            ref: "claudeMaxTokens",
                            label: "Max Tokens (response length)",
                            defaultValue: 2000,
                            min: 100,
                            max: 4000
                        },
                        testClaudeButton: {
                            component: "button",
                            label: "Test Claude AI Connection",
                            action: function(data) {
                                var mcpUrl = (data.mcpOrchestratorUrl || 'https://gse-mcp.replit.app').replace(/\/+$/, '');
                                var endpoint = '/api/execute-tool';
                                var fetchOptions = {
                                    method: 'POST',
                                    headers: { 'Content-Type': 'application/json' },
                                    credentials: 'include',
                                    body: JSON.stringify({
                                        tool: 'test-connection',
                                        args: { test: true }
                                    })
                                };
                                fetch(mcpUrl + endpoint, fetchOptions)
                                    .then(function(response) {
                                        if (!response.ok) {
                                            throw new Error('HTTP ' + response.status);
                                        }
                                        return response.json();
                                    })
                                    .then(function(result) {
                                        alert('✅ Claude AI Connection Successful!\n\nOrchestrator is reachable at:\n' + mcpUrl);
                                    })
                                    .catch(function(error) {
                                        alert('❌ Claude AI Connection Failed!\n\nError: ' + error.message + '\n\nCheck MCP Orchestrator URL');
                                    });
                            }
                        }
                    }
                },
                about: {
                    type: "items",
                    label: "About",
                    items: {
                        title: {
                            component: {
                                template: '<div style="text-align: center; padding: 10px 0;"><h1 style="color: #003B5C; font-size: 28px; font-weight: bold; margin: 0;">Template Qlik Extension</h1></div>'
                            }
                        },
                        version: {
                            component: {
                                template: '<div style="text-align: center; padding: 5px 0;"><span style="color: #555; font-size: 14px;">Version 1.0.0</span></div>'
                            }
                        },
                        spacer1: {
                            component: {
                                template: '<div style="text-align: center; padding: 15px 0;"><hr style="border: none; border-top: 1px solid #e0e0e0; width: 80%; margin: 0 auto;"></div>'
                            }
                        },
                        author: {
                            component: {
                                template: '<div style="text-align: center; padding: 5px 0;"><span style="color: #009845; font-size: 13px;">Created by: Jochem Zwienenberg</span></div>'
                            }
                        }
                    }
                },
                appearance: {
                    uses: "settings"
                }
            }
        },

        // Initial properties
        initialProperties: {
            qHyperCubeDef: {
                qDimensions: [],
                qMeasures: [],
                qInitialDataFetch: [{
                    qWidth: 10,
                    qHeight: 100
                }]
            }
        },

        // Paint method - called on every render
        paint: function ($element, layout) {
            // Clear element to avoid duplicate renders
            $element.empty();

            const props = layout.props || {};
            const exampleText = props.exampleText || "Hello from Template_Qlik!";
            const showBorder = props.showBorder || false;

            // Create container
            const container = $('<div class="template-qlik-container"></div>');

            if (showBorder) {
                container.addClass('with-border');
            }

            // Add example content
            const content = $('<div class="template-content"></div>');
            content.html(`<h2>${exampleText}</h2>`);

            // Check if we have data from hypercube
            if (layout.qHyperCube && layout.qHyperCube.qDataPages && layout.qHyperCube.qDataPages.length > 0) {
                const dataPage = layout.qHyperCube.qDataPages[0];

                if (dataPage.qMatrix && dataPage.qMatrix.length > 0) {
                    const dataList = $('<ul class="data-list"></ul>');

                    dataPage.qMatrix.forEach((row) => {
                        const listItem = $('<li></li>');
                        const rowText = row.map(cell => cell.qText).join(' - ');
                        listItem.text(rowText);
                        dataList.append(listItem);
                    });

                    content.append(dataList);
                } else {
                    content.append('<p class="no-data">No data available. Add dimensions or measures in the data panel.</p>');
                }
            } else {
                content.append('<p class="no-data">No data available. Add dimensions or measures in the data panel.</p>');
            }

            container.append(content);
            $element.append(container);

            return Promise.resolve();
        }
    };
});
