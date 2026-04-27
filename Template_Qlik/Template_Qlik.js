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
