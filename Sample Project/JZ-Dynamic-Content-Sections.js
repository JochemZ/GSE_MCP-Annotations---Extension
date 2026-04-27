// Version: 1.2.1 - Color Picker Fix + Background Color
// FIXES:
// - Text color picker now works correctly (selection preservation + proper focus)
// - Added background/highlight color picker
// - Improved color picker UI with visual indicators
// Version: 1.2.0 - ONE FILE PER PK (Major Architecture Improvement)
// BREAKING CHANGE: File structure changed from single file to one file per PK value
// - File naming: qlik-sections-{appId}-{pkValue}.json (was: qlik-sections-{appId}.json)
// - Data structure: Direct { sections: {...} } (was: { records: { pk: { sections: {...} } } })
// - Performance: Only loads/saves the specific PK file you're editing (much faster!)
// - Conflict prevention: Users editing different PKs CANNOT conflict (different files!)
// - Better locking: File-level locking now works perfectly (one user per file)
// - Scalability: Works with 1000s of PK values without performance degradation
// PREVIOUS FEATURES:
// - Consolidated multi-user settings: One MULTI_USER_ENABLED flag controls all collaboration features
// - File loading disabled when MULTI_USER_ENABLED = 0
// - Text formatting toolbar: Bold, Italic, Underline, Alignment, Font Size, Text Color, Lists
// - Undo button: Step back through previous changes
// - Reset button: Replace edited content with original template + live data values
// - Version-based conflict detection: Prevents simultaneous edit conflicts
// - Fast polling during edits: 5-second intervals when modal is open, prevents overwrites
// - View-first update flow: Extension view updates before modal opens, ensuring visual consistency
// - Real multi-user support: fetches authenticated user email from Qlik
// - Exponential backoff polling: reduces bandwidth by ~50% without orchestrator changes
// - Polling backs off from base interval (30s) to 2x, 4x, 8x when no changes detected
// - Resets to base interval immediately when changes detected (aggressive polling during activity)
// - Visual notifications when content is updated by others
// - Configurable base polling interval (10-300 seconds)
// - Clean console output (removed debug logs)

// DEV VERSION FLAG: Set to true for development builds (deployment script modifies this)
var IS_DEV_VERSION = false;

define([
    'jquery',
    'qlik',
    'require',
    'css!./style.css'
], function($, qlik, require) {
    'use strict';

    // Extension version (keep in sync with .qext file)
    var EXTENSION_VERSION = '2.0.25';

    // Get logo URL dynamically using require.toUrl for proper Qlik Cloud extension asset loading
    var QLIK_LOGO_URL = require.toUrl('./Img/Typemark_Qlik-Sales-Enablement-Color.png');

    // Store reference to current model for button actions
    var currentModel = null;

    // Track rendering state to prevent duplicate renders during async loading
    var renderingState = {};
    var currentRenderToken = {}; // Track current render to ignore stale callbacks
    var renderStartTime = {}; // Track when each render started
    var migrationDone = {}; // Track if migration has been completed for each instance
    var hasRenderedOnce = {}; // Track if instance has rendered at least once (to avoid spinner on properties open)

    // Track if save operation is in progress to prevent paint() during save
    var savingInProgress = false;

    // Prevent infinite render loops
    var paintCallCount = 0;
    var paintCallResetTimer = null;

    // Track last layout to detect property changes
    var lastLayoutHash = {};

    // Track SharePoint availability to prevent spam
    var sharepointUnavailable = false;
    var sharepointUnavailableSince = null;
    var sharepointErrorCount = {};  // Track errors per instance

    // Legacy variables (no longer used but kept for backward compatibility)
    var selectionListener = null;
    var lastSelectionState = null;
    var lastWarningMessage = null;
    var lastWarningTime = 0;

    // Track all created generic objects per instance for cleanup
    var genericObjectsRegistry = {};

    // Global paint counter - increments on every paint() call
    // Used to detect stale openSectionEditModal calls
    var globalPaintCounter = 0;

    // PK check token system - invalidates stale getCurrentPKValue callbacks
    var currentPKCheckToken = 0;
    var pkCheckObjects = [];
    var currentModalPKToken = null; // Track which token opened the current modal
    var currentModalPKObject = null; // Track generic object for current modal

    // Claude AI Feature Flag: Set to 1 to enable, 0 to disable completely
    var CLAUDE_ENABLED = 1;

    // Multi-User Feature Flag: Set to 1 to enable section editing, PK-based storage, user tracking, and polling
    // Set to 0 to disable all multi-user collaboration features completely
    var MULTI_USER_ENABLED = 1;

    // Polling interval in milliseconds (default: 30 seconds)
    var POLLING_INTERVAL = 30000;

    // Current authenticated user info (fetched once at load)
    var currentUserEmail = 'user@qlik.com'; // Default fallback
    var currentUserName = 'Unknown User'; // Default fallback
    var currentUserData = null; // Full user object

    // User feature permissions (fetched from orchestrator based on SSO email)
    var userAllowedFeatures = []; // List of feature keys user has permission to access
    var userPermissionsLoaded = false; // Track if permissions have been fetched
    var userPermissionsFailed = false; // Track if permissions fetch failed

    // Polling state
    var pollingTimers = {}; // One timer per extension instance
    var lastKnownData = {}; // Track last known data per instance to detect changes
    var pollingIntervals = {}; // Current interval per instance (for exponential backoff)
    var consecutiveNoChanges = {}; // Track consecutive polls with no changes

    // Modal state for fast polling and conflict detection
    var modalIsOpen = false; // Track if edit modal is currently open
    var modalBaselineTimestamp = null; // Baseline timestamp when modal was opened
    var modalSectionLabel = null; // Which section is being edited
    var modalPKValue = null; // Which PK value is being edited
    var FAST_POLLING_INTERVAL = 5000; // 5 seconds when modal is open

    // Cache for Claude AI responses (keyed by data hash)
    var claudeResponseCache = {};

    // Cache for modified sections (loaded from file)
    var modifiedSectionsCache = {};
    var modifiedSectionsLastCheck = 0;
    var POLLING_INTERVAL = 60000; // 60 seconds
    var RECENT_EDIT_THRESHOLD = 10 * 60 * 1000; // 10 minutes

    // Cache for master dimensions - populated once, used synchronously
    var masterDimensionsCache = [{ value: "", label: "Loading dimensions..." }];
    var masterDimensionsCached = false;

    // Cache for master items (dimensions + measures) to avoid re-fetching on every paint
    var masterItemsCache = {
        dimensions: null,
        measures: null,
        cached: false
    };

    // Default Qlik color palette (defined early to ensure availability)
    var defaultPalette = [
        '#009845', '#1976D2', '#EE9A00', '#C62828', '#6B1B7F',
        '#00ACC1', '#7CB342', '#FF7043', '#AB47BC', '#26A69A'
    ];

    // Simple hash function for data
    function hashData(data) {
        return JSON.stringify(data);
    }

    // Update loading message dynamically
    function updateLoadingMessage($element, mainMessage, subMessage) {
        var $spinner = $element.find('.jz-loading-spinner');
        if ($spinner.length === 0) return;

        var $text = $spinner.find('.jz-loading-spinner-text');
        var $subtext = $spinner.find('.jz-loading-spinner-subtext');

        if ($text.length > 0) {
            $text.text(mainMessage);
        }

        if (subMessage) {
            if ($subtext.length === 0) {
                $spinner.append('<div class="jz-loading-spinner-subtext">' + subMessage + '</div>');
            } else {
                $subtext.text(subMessage);
            }
        } else {
            $subtext.remove();
        }
    }

    // Load color dimensions from markdown and populate colorMappings
    function loadColorDimensionsFromMarkdown(layout, context) {
        var app = qlik.currApp(context);

        // Extract colorBy dimensions from all sections in all groups
        var colorByDimensions = new Set();
        var groups = layout.groups || [];

        // Handle backward compatibility
        if (groups.length === 0 && layout.sections && layout.sections.length > 0) {
            groups = [{sections: layout.sections}];
        }

        groups.forEach(function(group) {
            if (group.sections) {
                group.sections.forEach(function(section) {
                    if (section.markdownText) {
                        var colorByPattern = /colorBy="\[([^\]]+)\]"/g;
                        var match;
                        while ((match = colorByPattern.exec(section.markdownText)) !== null) {
                            colorByDimensions.add(match[1]);
                        }
                    }
                });
            }
        });

        if (colorByDimensions.size === 0) {
            alert('No colorBy dimensions found in your markdown.\n\nAdd colorBy="[Dimension Name]" to your content tags first.\n\nExample: #[list colorBy="[Strategic Player Role]"]{{[Strategic Player]}}#[/list]');
            return;
        }

        var dimensionsArray = Array.from(colorByDimensions);
        var allColorMappings = layout.colorMappings || [];

        // Get master dimensions list
        app.createGenericObject({
            qInfo: { qType: 'DimensionList' },
            qDimensionListDef: { qType: 'dimension', qData: { title: '/qMetaDef/title' } }
        }, function(dimensionReply) {
            var dimensionList = dimensionReply.qDimensionList.qItems;
            var dimensionMap = {};

            dimensionList.forEach(function(dim) {
                dimensionMap[dim.qMeta.title] = dim;
            });

            var totalDimensions = dimensionsArray.length;
            var processedCount = 0;
            var totalNew = 0;

            // For each colorBy dimension, create a hypercube to get distinct values
            dimensionsArray.forEach(function(dimName) {
                var dimItem = dimensionMap[dimName];
                if (!dimItem) {
                    
                    processedCount++;
                    checkIfComplete();
                    return;
                }

                // Create hypercube with this dimension
                app.createCube({
                    qDimensions: [{
                        qLibraryId: dimItem.qInfo.qId,
                        qDef: {
                            qFieldDefs: [],
                            qSortCriterias: [{ qSortByAscii: 1 }]
                        }
                    }],
                    qMeasures: [],
                    qInitialDataFetch: [{
                        qTop: 0,
                        qLeft: 0,
                        qHeight: 100,
                        qWidth: 1
                    }]
                }, function(reply) {

                    // Extract distinct values
                    if (reply.qHyperCube && reply.qHyperCube.qDataPages[0]) {
                        var rows = reply.qHyperCube.qDataPages[0].qMatrix;

                        rows.forEach(function(row) {
                            var value = row[0].qText;

                            // Check if mapping already exists
                            var existingMapping = allColorMappings.find(function(m) {
                        return m.dimensionValue === value;
                            });

                            if (!existingMapping) {
                                // Generate default color using palette for text
                        var defaultColor = getDefaultPaletteColor(value);
                                // Use white background by default (users can customize)
                        var defaultBgColor = '#FFFFFF';

                        allColorMappings.push({
                                    dimensionName: dimName,
                                    dimensionValue: value,
                                    textColor: defaultColor,
                                    bgColor: defaultBgColor
                                });
                        totalNew++;
                            }
                        });

                        // 
                    }

                    processedCount++;
                    checkIfComplete();
                });
            });

            function checkIfComplete() {
                if (processedCount === totalDimensions) {
                    // Update the layout
                    layout.colorMappings = allColorMappings;

                    alert('Loaded ' + totalNew + ' new dimension values.\n\nTotal mappings: ' + allColorMappings.length + '\n\nScroll down in "Custom Dimension Value Colors" section to assign colors using the color pickers.');

                    // Force property panel refresh
                    if (context && context.resize) {
                        context.resize();
                    }
                }
            }
        });
    }

    // Color configuration modal
    function showColorConfigModal(dimensionValues, existingColors, onSave) {
        var colorMapping = $.extend({}, existingColors);

        var $overlay = $('<div class="edit-modal-overlay"></div>');
        var $modal = $('<div class="edit-modal"></div>');

        var $header = $('<div class="modal-header"></div>');
        $header.append('<h3>Configure Dimension Colors</h3>');

        var $closeBtn = $('<button class="modal-close-btn">&times;</button>').on('click', function() {
            $overlay.remove();
        });
        $header.append($closeBtn);

        var $body = $('<div class="modal-body"></div>');

        dimensionValues.forEach(function(value) {
            var currentColor = colorMapping[value] || '#333333';

            var $row = $('<div style="display: flex; align-items: center; gap: 12px; margin-bottom: 12px;"></div>');
            var $label = $('<label style="flex: 1; font-size: 14px;"></label>').text(value);
            var $colorInput = $('<input type="color" style="width: 60px; height: 32px; cursor: pointer;">').val(currentColor);

            $colorInput.on('change', function() {
                colorMapping[value] = $(this).val();
            });

            $row.append($label, $colorInput);
            $body.append($row);
        });

        var $footer = $('<div class="modal-footer"></div>');
        var $cancelBtn = $('<button class="modal-btn cancel-btn">Cancel</button>').on('click', function() {
            $overlay.remove();
        });
        var $saveBtn = $('<button class="modal-btn save-btn">Save Colors</button>').on('click', function() {
            onSave(colorMapping);
            $overlay.remove();
        });

        $footer.append($cancelBtn, $saveBtn);
        $modal.append($header, $body, $footer);
        $overlay.append($modal);
        $('body').append($overlay);

        $overlay.on('click', function(e) {
            if (e.target === $overlay[0]) {
                $overlay.remove();
            }
        });
    }

    return {
        initialProperties: {
            qHyperCubeDef: {
                qDimensions: [],
                qMeasures: [],
                qInitialDataFetch: [{
                    qTop: 0,
                    qLeft: 0,
                    qHeight: 100,
                    qWidth: 20
                }],
                qMode: "S",
                qSuppressZero: false,
                qSuppressMissing: false,
                qAlwaysFullyExpanded: true
            },
            groups: [{
                groupLabel: "Default Group",
                groupWidth: "full",
                groupBgColor: "transparent",
                sections: [{
                    label: "Welcome Section",
                    markdownText: "# Welcome to Dynamic Content Sections!\n\nAdd dimensions and measures in the Data panel, then reference them using:\n- Master items: `{{[Master Item Name]}}`\n- Direct reference: `{{dim1}}`, `{{measure1}}`\n\nUse content tags like:\n- `#[list]{{dim1}}#[/list]`\n- `#[table]Product|{{dim1}}\\nRevenue|{{measure1}}#[/table]`\n- `#[kpi label=\"Total\"]{{measure1}}#[/kpi]`\n\nClick the example buttons below to see more!",
                    hideIfNoData: false,
                    sectionStyle: "card",
                    sectionWidth: "full",
                    sectionBgColor: "transparent",
                    enableEdit: false,
                    showLabel: false,
                    labelSeparator: false,
                    labelStyle: "bold",
                    labelColor: "#1a1a1a",
                    iconAction: "none",
                    iconLink: ""
                }]
            }],
            spacing: 5,
            padding: 3,
            colorMapsList: []
        },

        support: {
            snapshot: true,
            export: true,
            exportData: false
        },

        definition: {
            type: "items",
            component: "accordion",
            items: {
                groups: {
                    type: "array",
                    ref: "groups",
                    label: "Groups",
                    itemTitleRef: "groupLabel",
                    allowAdd: true,
                    allowRemove: true,
                    allowMove: true,
                    addTranslation: "Add Group",
                    items: {
                        groupLabel: {
                            type: "string",
                            ref: "groupLabel",
                            label: "Group Label",
                            defaultValue: "New Group"
                        },
                        groupWidth: {
                            type: "string",
                            component: "dropdown",
                            label: "Group Width",
                            ref: "groupWidth",
                            options: [{
                        value: "full",
                        label: "Full Width (1/1)"
                            }, {
                        value: "half",
                        label: "Half Width (1/2)"
                            }, {
                        value: "third",
                        label: "Third Width (1/3)"
                            }, {
                        value: "quarter",
                        label: "Quarter Width (1/4)"
                            }],
                            defaultValue: "full"
                        },
                        groupBgColor: {
                            type: "string",
                            ref: "groupBgColor",
                            label: "Group Background Color (hex code or 'transparent')",
                            defaultValue: "transparent",
                            expression: "optional"
                        },
                        groupSpacing: {
                            type: "number",
                            ref: "groupSpacing",
                            label: "Group Spacing (px) - overrides global spacing for this group",
                            defaultValue: null,
                            expression: "optional"
                        },
                        groupBorderColor: {
                            type: "string",
                            ref: "groupBorderColor",
                            label: "Group Border Color (hex code or 'transparent')",
                            defaultValue: "transparent",
                            expression: "optional"
                        },
                        groupBorderWidth: {
                            type: "number",
                            ref: "groupBorderWidth",
                            label: "Group Border Width (px)",
                            defaultValue: 0,
                            expression: "optional"
                        },
                        groupBorderStyle: {
                            type: "string",
                            component: "dropdown",
                            ref: "groupBorderStyle",
                            label: "Group Border Style",
                            options: [{
                                value: "solid",
                                label: "Solid"
                            }, {
                                value: "dashed",
                                label: "Dashed"
                            }, {
                                value: "dotted",
                                label: "Dotted"
                            }],
                            defaultValue: "solid",
                            expression: "optional"
                        },
                        sections: {
                            type: "array",
                            ref: "sections",
                            label: "Sections",
                            itemTitleRef: "label",
                            allowAdd: true,
                            allowRemove: true,
                            allowMove: true,
                            addTranslation: "Add Section",
                            items: {
                        label: {
                                    type: "string",
                                    ref: "label",
                                    label: "Section Label",
                                    defaultValue: "New Section"
                                },
                        markdownText: {
                                    type: "string",
                                    component: "textarea",
                                    label: "Content (use {{[Master Item Name]}}, e.g. {{[Strategic Player]}}, and tags like #[header], #[title], #[table], #[list], #[kpi], etc.)",
                                    ref: "markdownText",
                                    defaultValue: "# Section Title\n\nUse master items like:\n{{[Strategic Player]}}\n\nOr in tags:\n#[list]{{[Product]}}#[/list]\n\nColors: {red:red text}, {blue:blue text}, {#009845:custom hex}",
                                    rows: 20,
                                    maxlength: 10000
                                },
                        guiEditorButton: {
                                    component: "button",
                                    label: "📝 Open Visual Content Editor (Beta)",
                                    show: function() {
                                        return userHasPropertyPermission('jz_Properties_Visual_Design');
                                    },
                                    action: function(data, args, layout) {
                                        // data = the current section object
                                        // args = handler object with properties/model
                                        // layout = full extension layout

                                        // console.log('[GUI EDITOR] Action called with args:', args);
                                        // console.log('[GUI EDITOR] Has handler?', !!args);
                                        if (args) {
                                            // console.log('[GUI EDITOR] Handler keys:', Object.keys(args));
                                        }

                                        var section = data;

                                        var dummyLayout = {
                                            enableClaude: false
                                        };

                                        // Open GUI editor
                                        showGUIEditorModal(section, null, dummyLayout, function(newContent) {
                                            // console.log('[GUI EDITOR] Content saved:', newContent);

                                            // CRITICAL: Update args.properties, not just the local section reference
                                            // Find this section in args.properties.groups and update it there
                                            if (args && args.properties && args.properties.groups) {
                                                var sectionFound = false;

                                                for (var gi = 0; gi < args.properties.groups.length; gi++) {
                                                    var group = args.properties.groups[gi];
                                                    if (group.sections) {
                                                        for (var si = 0; si < group.sections.length; si++) {
                                                            if (group.sections[si].label === section.label) {
                                                                // Found it! Update the actual properties object
                                                                // console.log('[GUI EDITOR] Found section in properties at group', gi, 'section', si);
                                                                args.properties.groups[gi].sections[si].markdownText = newContent;
                                                                sectionFound = true;
                                                                break;
                                                            }
                                                        }
                                                    }
                                                    if (sectionFound) break;
                                                }

                                                if (sectionFound) {
                                                    // console.log('[GUI EDITOR] ✓ Updated args.properties - now persisting with setProperties()...');

                                                    // ========================================================================
                                                    // CRITICAL FIX - DO NOT REMOVE OR MODIFY (Fixed: 2026-03-19)
                                                    // ========================================================================
                                                    // PROBLEM: WYSIWYG editor changes were lost when leaving edit mode
                                                    //   - Updating args.properties alone doesn't persist to Qlik's model
                                                    //   - Toggle hacks (hideIfNoData flip/flop) are unreliable
                                                    //   - Properties panel showed updates inconsistently
                                                    //   - Changes disappeared when returning to sheet view
                                                    //
                                                    // SOLUTION: Use proper Qlik API to persist changes
                                                    //   1. Get visualization object: app.visualization.get(objectId)
                                                    //   2. Get fresh properties: vis.model.getProperties()
                                                    //   3. Update fresh copy with new content
                                                    //   4. Persist using: vis.model.setProperties(props)
                                                    //
                                                    // This ensures:
                                                    //   ✓ Changes persist when leaving/returning to edit mode
                                                    //   ✓ Properties panel updates reliably
                                                    //   ✓ Qlik's autosave is properly triggered
                                                    //
                                                    // REFERENCE: Same pattern as "Apply Changes to Sheet" button (removed)
                                                    // ========================================================================
                                                    if (args.layout && args.layout.qInfo && args.layout.qInfo.qId) {
                                                        var objectId = args.layout.qInfo.qId;
                                                        var app = qlik.currApp();
                                                        var savedGroupIndex = gi;
                                                        var savedSectionIndex = si;

                                                        app.visualization.get(objectId).then(function(vis) {
                                                            // console.log('[GUI EDITOR] Got visualization object');

                                                            // Get current properties and update with our changes
                                                            vis.model.getProperties().then(function(props) {
                                                                // console.log('[GUI EDITOR] Got fresh properties from model');

                                                                // Update the fresh properties copy with the new content
                                                                if (props.groups && props.groups[savedGroupIndex] &&
                                                                    props.groups[savedGroupIndex].sections &&
                                                                    props.groups[savedGroupIndex].sections[savedSectionIndex]) {

                                                                    props.groups[savedGroupIndex].sections[savedSectionIndex].markdownText = newContent;
                                                                    // console.log('[GUI EDITOR] Updated fresh props with new content');

                                                                    // Now persist the updated properties
                                                                    vis.model.setProperties(props).then(function() {
                                                                        // console.log('[GUI EDITOR] ✓ setProperties() succeeded - changes persisted!');
                                                                    }).catch(function(err) {
                                                                        // console.error('[GUI EDITOR] Error in setProperties:', err);
                                                                    });
                                                                } else {
                                                                    // console.error('[GUI EDITOR] Could not find section in fresh properties');
                                                                }
                                                            }).catch(function(err) {
                                                                // console.error('[GUI EDITOR] Error getting properties:', err);
                                                            });
                                                        }).catch(function(err) {
                                                            // console.error('[GUI EDITOR] Error getting visualization:', err);
                                                        });
                                                    } else {
                                                        // console.warn('[GUI EDITOR] No qInfo.qId - cannot persist changes');
                                                    }
                                                } else {
                                                    // console.log('[GUI EDITOR] ⚠ Could not find section in properties');
                                                }
                                            }

                                            // Also update the local reference
                                            section.markdownText = newContent;

                                            // Clear SharePoint cache
                                            if (modifiedSectionsCache.data && modifiedSectionsCache.data.sections) {
                                                if (modifiedSectionsCache.data.sections[section.label]) {
                                                    // console.log('[GUI EDITOR] Clearing SharePoint cache');
                                                    delete modifiedSectionsCache.data.sections[section.label];
                                                }
                                            }

                                            // console.log('[GUI EDITOR] ✓ Save complete - sheet should repaint shortly!');
                                        });
                                    }
                                },
                        hideIfNoData: {
                                    type: "boolean",
                                    ref: "hideIfNoData",
                                    label: "Hide section if no data",
                                    defaultValue: false,
                                    change: function(data) {
                                        // console.log('[CHECKBOX] hideIfNoData changed to:', data.hideIfNoData);
                                        // console.log('[CHECKBOX] This change will trigger paint() automatically');
                                        // console.log('[CHECKBOX] No explicit save() call needed - Qlik handles it via "ref"');
                                    }
                                },
                        sectionStyle: {
                                    type: "string",
                                    component: "dropdown",
                                    label: "Section Style",
                                    ref: "sectionStyle",
                                    options: [{
                                        value: "card",
                                        label: "Card (with border)"
                                    }, {
                                        value: "plain",
                                        label: "Plain (no border)"
                                    }, {
                                        value: "highlighted",
                                        label: "Highlighted"
                                    }],
                                    defaultValue: "card"
                                },
                        sectionWidth: {
                                    type: "string",
                                    component: "dropdown",
                                    label: "Section Width",
                                    ref: "sectionWidth",
                                    options: [{
                                        value: "full",
                                        label: "Full Width (1/1)"
                                    }, {
                                        value: "half",
                                        label: "Half Width (1/2)"
                                    }, {
                                        value: "third",
                                        label: "Third Width (1/3)"
                                    }, {
                                        value: "quarter",
                                        label: "Quarter Width (1/4)"
                                    }],
                                    defaultValue: "full"
                                },
                        sectionBgColor: {
                                    type: "string",
                                    ref: "sectionBgColor",
                                    label: "Section Background Color (hex code or 'transparent')",
                                    defaultValue: "transparent",
                                    expression: "optional"
                                },
                        enableEdit: {
                                    type: "boolean",
                                    ref: "enableEdit",
                                    label: "Enable editing for this section",
                                    defaultValue: false,
                                    show: function() {
                                        return MULTI_USER_ENABLED === 1 && userHasPropertyPermission('jz_edit');
                                    }
                                },
                        showLabel: {
                                    type: "boolean",
                                    ref: "showLabel",
                                    label: "Show label as title",
                                    defaultValue: false
                                },
                        labelSeparator: {
                                    type: "boolean",
                                    ref: "labelSeparator",
                                    label: "Label separator line",
                                    defaultValue: false
                                },
                        labelStyle: {
                                    type: "string",
                                    component: "dropdown",
                                    label: "Label style",
                                    ref: "labelStyle",
                                    options: [{
                                        value: "normal",
                                        label: "Normal"
                                    }, {
                                        value: "bold",
                                        label: "Bold"
                                    }, {
                                        value: "italic",
                                        label: "Italic"
                                    }],
                                    defaultValue: "bold"
                                },
                        labelColor: {
                                    type: "string",
                                    ref: "labelColor",
                                    label: "Label color (hex code or color name)",
                                    defaultValue: "#1a1a1a"
                                },
                        iconAction: {
                                    type: "string",
                                    component: "dropdown",
                                    label: "Icon Action",
                                    ref: "iconAction",
                                    options: [{
                                        value: "none",
                                        label: "No Icon"
                                    }, {
                                        value: "clearSelections",
                                        label: "Clear All Selections"
                                    }, {
                                        value: "homeLink",
                                        label: "Home Link (Internal)"
                                    }, {
                                        value: "customLink",
                                        label: "Custom Link"
                                    }],
                                    defaultValue: "none"
                                },
                        iconLink: {
                                    type: "string",
                                    ref: "iconLink",
                                    label: "Custom Link URL",
                                    defaultValue: ""
                                },
                        applyColorMap: {
                                    type: "string",
                                    ref: "applyColorMap",
                                    label: "Apply Color Map (auto-colors field values)",
                                    defaultValue: "none",
                                    component: "dropdown",
                                    options: function(data, handler) {
                                        var options = [{ value: "none", label: "None (no automatic coloring)" }];
                                        if (data.colorMapsList && data.colorMapsList.length > 0) {
                                            data.colorMapsList.forEach(function(map) {
                                                options.push({
                                                    value: map.mapName || "Unnamed",
                                                    label: map.mapName || "Unnamed"
                                                });
                                            });
                                        }
                                        return options;
                                    }
                                },
                        colorMapField: {
                                    type: "string",
                                    ref: "colorMapField",
                                    label: "Apply to Field (e.g., [Revenue Aspirations])",
                                    defaultValue: "",
                                    expression: "optional",
                                    show: function(data) {
                                        return data && data.applyColorMap && data.applyColorMap !== "none";
                                    }
                                },
                        colorMapExtractPattern: {
                                    type: "string",
                                    ref: "colorMapExtractPattern",
                                    label: "Extract Pattern (regex to get value, e.g., \\(([a-zA-Z0-9]+)\\) for Last Period (1) or (z))",
                                    defaultValue: "\\(([a-zA-Z0-9]+)\\)",
                                    expression: "optional",
                                    show: function(data) {
                                        return data && data.applyColorMap && data.applyColorMap !== "none";
                                    }
                                },
                        colorMapLabel: {
                                    type: "string",
                                    ref: "colorMapLabel",
                                    label: "Label Prefix (e.g., 'Last Period' to show before value)",
                                    defaultValue: "",
                                    expression: "optional",
                                    show: function(data) {
                                        return data && data.applyColorMap && data.applyColorMap !== "none";
                                    }
                                }
                            }
                        }
                    }
                },
                dimensionValueColors: {
                    type: "items",
                    label: "Custom Dimension Value Colors",
                    items: {
                        colorMappingsInfo: {
                            component: "text",
                            label: "Click 'Load Color Dimensions' to automatically detect colorBy dimensions from your markdown and load their values."
                        },
                        loadColorDimensionsButton: {
                            component: "button",
                            label: "Load Color Dimensions from Markdown",
                            action: function(data) {
                        loadColorDimensionsFromMarkdown(data, this);
                            }
                        },
                        colorMappings: {
                            type: "array",
                            ref: "colorMappings",
                            label: "Dimension Value Colors (click to edit)",
                            itemTitleRef: function(item) {
                        if (item.dimensionName && item.dimensionValue) {
                                    return '[' + item.dimensionName + '] ' + item.dimensionValue;
                                } else if (item.dimensionValue) {
                                    return item.dimensionValue;
                                }
                        return 'New mapping';
                            },
                            allowAdd: true,
                            allowRemove: true,
                            allowMove: false,
                            addTranslation: "Add Manual Mapping",
                            items: {
                        dimensionName: {
                                    type: "string",
                                    ref: "dimensionName",
                                    label: "Dimension Name",
                                    expression: "optional",
                                    defaultValue: "",
                                    show: false  // Hidden since it's auto-populated
                                },
                        dimensionValue: {
                                    type: "string",
                                    ref: "dimensionValue",
                                    label: "Dimension Value",
                                    expression: "optional",
                                    defaultValue: ""
                                },
                        textColor: {
                                    type: "string",
                                    ref: "textColor",
                                    label: "Text Color",
                                    component: "color-picker",
                                    defaultValue: "#000000"
                                },
                        bgColor: {
                                    type: "string",
                                    ref: "bgColor",
                                    label: "Background Color",
                                    component: "color-picker",
                                    defaultValue: "#FFFFFF"
                                }
                            }
                        }
                    }
                },
                appearance: {
                    uses: "settings",
                    items: {
                        spacing: {
                            type: "number",
                            ref: "spacing",
                            label: "Spacing between sections (px)",
                            defaultValue: 5,
                            min: 0,
                            max: 50
                        },
                        padding: {
                            type: "number",
                            ref: "padding",
                            label: "Section padding (px)",
                            defaultValue: 6,
                            min: 0,
                            max: 50
                        },
                        fontFamily: {
                            type: "string",
                            component: "dropdown",
                            ref: "fontFamily",
                            label: "Font Family",
                            options: [
                                { value: "'QlikView Sans', sans-serif", label: "QlikView Sans" },
                                { value: "'Source Sans Pro', sans-serif", label: "Source Sans Pro" },
                                { value: "Arial, sans-serif", label: "Arial" },
                                { value: "'Helvetica Neue', Helvetica, sans-serif", label: "Helvetica" },
                                { value: "'Open Sans', sans-serif", label: "Open Sans" },
                                { value: "Inter, sans-serif", label: "Inter" },
                                { value: "Roboto, sans-serif", label: "Roboto" },
                                { value: "'Segoe UI', sans-serif", label: "Segoe UI" },
                                { value: "Georgia, serif", label: "Georgia" },
                                { value: "'Times New Roman', serif", label: "Times New Roman" },
                                { value: "'Courier New', monospace", label: "Courier New" }
                            ],
                            defaultValue: "'Source Sans Pro', sans-serif"
                        },
                        fontSize: {
                            type: "string",
                            component: "dropdown",
                            ref: "fontSize",
                            label: "Font Size",
                            options: [
                                { value: "10", label: "10px" },
                                { value: "11", label: "11px" },
                                { value: "12", label: "12px" },
                                { value: "13", label: "13px" },
                                { value: "14", label: "14px" },
                                { value: "15", label: "15px" },
                                { value: "16", label: "16px" },
                                { value: "18", label: "18px" },
                                { value: "20", label: "20px" },
                                { value: "22", label: "22px" },
                                { value: "24", label: "24px" }
                            ],
                            defaultValue: "14"
                        },
                        labelSizeOffset: {
                            type: "number",
                            component: "slider",
                            ref: "labelSizeOffset",
                            label: "Section label size offset (px added to default font size)",
                            min: -4,
                            max: 12,
                            step: 1,
                            defaultValue: 4
                        },
                        _refreshTrigger: {
                            type: "number",
                            ref: "_refreshTrigger",
                            defaultValue: 0,
                            show: false
                        }
                    }
                },
                colorMaps: {
                    type: "items",
                    label: "Color Maps",
                    items: {
                        colorMapsInfo: {
                            component: "text",
                            label: "Define reusable color mappings for automatic field coloring. Format: value1:textColor1,value2:textColor2 OR value1:textColor1:bgColor1,value2:textColor2:bgColor2"
                        },
                        colorMapsList: {
                            type: "array",
                            ref: "colorMapsList",
                            label: "Color Maps",
                            itemTitleRef: "mapName",
                            allowAdd: true,
                            allowRemove: true,
                            allowMove: true,
                            addTranslation: "Add Color Map",
                            items: {
                        mapName: {
                                    type: "string",
                                    ref: "mapName",
                                    label: "Map Name",
                                    defaultValue: "New Color Map",
                                    expression: "optional"
                                },
                        mapDefinition: {
                                    type: "string",
                                    component: "textarea",
                                    ref: "mapDefinition",
                                    label: "Color Mapping (value:textColor or value:textColor:bgColor)",
                                    defaultValue: "1:#C62828:#FFCDD2,2:#E65100:#FFE0B2,3:#F9A825:#FFF9C4,4:#43A047:#C8E6C9,5:#1B5E20:#A5D6A7",
                                    rows: 3,
                                    expression: "optional"
                                },
                        mapDescription: {
                                    type: "string",
                                    ref: "mapDescription",
                                    label: "Description (optional)",
                                    defaultValue: "",
                                    expression: "optional"
                                }
                            }
                        }
                    }
                },
                mcpOrchestratorSettings: {
                    type: "items",
                    label: "GSE - MCP Orchestrator Settings",
                    show: function() {
                        return (MULTI_USER_ENABLED === 1 || CLAUDE_ENABLED === 1) && userHasPropertyPermission('jz_Properties_GSE_MCP_Orchestrator');
                    },
                    items: {
                        mcpHeader: {
                            component: "text",
                            label: "Configure MCP Orchestrator URL for integration with SharePoint and Claude AI."
                        },
                        generalSettingsGroup: {
                            type: "items",
                            label: "⚙️ General Settings",
                            items: {
                                mcpOrchestratorUrl: {
                                    type: "string",
                                    ref: "mcpOrchestratorUrl",
                                    label: "MCP Orchestrator URL (no trailing slash)",
                                    defaultValue: "https://gse-mcp.replit.app",
                                    expression: "optional",
                                    show: function() {
                                        return MULTI_USER_ENABLED === 1 || CLAUDE_ENABLED === 1;
                                    }
                                },
                                ssoInfo: {
                                    component: "text",
                                    label: "🔒 Authentication: SSO (session-based via Qlik Cloud)",
                                    show: function() {
                                        return MULTI_USER_ENABLED === 1 || CLAUDE_ENABLED === 1;
                                    }
                                }
                            }
                        },
                    }
                },
                mcpSharePointSettings: {
                    type: "items",
                    label: "GSE - MCP SharePoint",
                    show: function() {
                        return MULTI_USER_ENABLED === 1 && userHasPropertyPermission('jz_properties_GSE_MCP_SharePoint');
                    },
                    items: {
                        sharepointHeader: {
                            component: "text",
                            label: "Configure SharePoint storage for multi-user collaboration."
                        },
                        sharepointSitePath: {
                            type: "string",
                            ref: "sharepointSitePath",
                            label: "SharePoint Site Path (e.g., /sites/QlikTechnicalEnablement)",
                            defaultValue: "/sites/QlikTechnicalEnablement",
                            expression: "optional"
                        },
                        sharepointFolderPath: {
                            type: "string",
                            ref: "sharepointFolderPath",
                            label: "SharePoint Folder Path (e.g., /Shared Documents/Qlik200_Topsheet)",
                            defaultValue: "/Shared Documents/Qlik200_Topsheet",
                            expression: "optional"
                        },
                        testSharePointButton: {
                            component: "button",
                            label: "Test SharePoint Connection",
                            action: function(data) {
                        var mcpUrl = (data.mcpOrchestratorUrl || 'http://localhost:5000').replace(/\/+$/, '');
                        var sitePath = data.sharepointSitePath || '/sites/YourSiteName';
                        var folderPath = data.sharepointFolderPath || '/Documents/QlikExtensions';

                        var headers = { 'Content-Type': 'application/json' };

                                // Build fetch options with SSO credentials
                        var fetchOptions = {
                                    method: 'POST',
                                    headers: headers,
                                    credentials: 'include', // SSO mode - always include cookies
                                    body: JSON.stringify({
                                        sitePath: sitePath,
                                        folderPath: folderPath
                                    })
                                };

                        fetch(mcpUrl + '/api/list-json-sharepoint', fetchOptions)
                                .then(function(response) {
                                    if (!response.ok) {
                                        return response.text().then(function(text) {
                                            throw new Error('HTTP ' + response.status + ': ' + text);
                                        });
                                    }
                                    return response.json();
                                })
                                .then(function(result) {
                                    if (result.success) {
                                        var fileCount = result.files ? result.files.length : 0;
                                        alert('✅ SharePoint Connection Successful!\n\n' +
                                              'Found ' + fileCount + ' file(s) in:\n' +
                                              sitePath + folderPath);
                                    } else {
                                        alert('⚠️ SharePoint connection succeeded but returned no data.\n\n' +
                                              'The folder might be empty or not exist yet.');
                                        
                                    }
                                })
                                .catch(function(error) {
                                    alert('❌ SharePoint Connection Failed!\n\n' +
                                          'Error: ' + error.message + '\n\n' +
                                          'Please check:\n' +
                                          '1. MCP Orchestrator is running on ' + mcpUrl + '\n' +
                                          '2. Site Path and Folder Path are correct\n' +
                                          '3. Orchestrator CORS is configured for SSO');
                                    
                                });
                            }
                        }
                    }
                },
                mcpClaudeSettings: {
                    type: "items",
                    label: "GSE - MCP Claude",
                    show: function() {
                        return CLAUDE_ENABLED === 1 && userHasPropertyPermission('jz_properties_GSE_MCP_Claude');
                    },
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
                            defaultValue: "You are a business intelligence analyst. Answer directly with only the requested information. Do not include preambles like 'Here are' or 'Based on the data'. Just provide the answer.",
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
                                var headers = { 'Content-Type': 'application/json' };
                                var fetchOptions = {
                                    method: 'POST',
                                    headers: headers,
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
                        },
                        clearCacheButton: {
                            component: "button",
                            label: "Clear Claude Cache & Refresh",
                            action: function(data) {
                                var cacheSize = Object.keys(claudeResponseCache).length;
                                claudeResponseCache = {};
                                alert('✅ Cache Cleared!\n\n' + cacheSize + ' cached response(s) removed.\n\nClaude AI will analyze data fresh on next load.');
                            }
                        }
                    }
                },
                multiUserSettings: {
                    type: "items",
                    label: "GSE - Multi-User",
                    show: function() {
                        return MULTI_USER_ENABLED === 1 && userHasPropertyPermission('jz_properties_GSE_MCP_SharePoint');
                    },
                    items: {
                        multiUserInfo: {
                            component: "text",
                            label: "Enable section editing, PK-based storage, user tracking, and real-time collaboration with automatic polling."
                        },
                        pkField: {
                            type: "string",
                            ref: "pkField",
                            label: "Primary Key Field (e.g., AccountID, CustomerName)",
                            defaultValue: "",
                            expression: "optional"
                        },
                        pkFieldInfo: {
                            component: "text",
                            label: "⚠️ Important: Select exactly ONE value in this field to enable editing. Different selections will load different saved content."
                        },
                        currentUser: {
                            component: "text",
                            label: function() {
                        if (currentUserName !== 'Unknown User' && currentUserEmail !== 'user@qlik.com') {
                                    return "Current User: " + currentUserName + " (" + currentUserEmail + ")";
                                } else if (currentUserEmail !== 'user@qlik.com') {
                                    return "Current User: " + currentUserEmail;
                                } else {
                                    return "Current User: Loading...";
                                }
                            }
                        },
                        pollingInterval: {
                            type: "number",
                            ref: "pollingInterval",
                            label: "Polling Interval (seconds)",
                            defaultValue: 30,
                            min: 10,
                            max: 300,
                            expression: "optional"
                        },
                        pollingInfo: {
                            component: "text",
                            label: "ℹ️ Uses exponential backoff: polls at base interval when changes detected, then backs off to 2x, 4x, 8x if no changes. Reduces bandwidth by ~50%."
                        },
                        pollingExample: {
                            component: "text",
                            label: "Example with 30s base: 30s → 30s → 60s → 120s → 240s (max 240s). Resets to 30s when changes detected."
                        }
                    }
                },
                importExport: {
                    type: "items",
                    label: "Import / Export Settings",
                    items: {
                        importExportInfo: {
                            component: "text",
                            label: "Backup and restore all extension settings (groups, sections, colors, color maps, etc.)"
                        },
                        exportButton: {
                            component: "button",
                            label: "Export Settings to File",
                            action: function(data) {
                                try {
                                    // Properties to EXCLUDE from export (Qlik internal + app-specific)
                                    var excludeProps = [
                                        'qInfo',
                                        'qMetaDef',
                                        'qHyperCubeDef',
                                        'qListObjectDef',
                                        'qDef',
                                        'visualization',
                                        'version',
                                        'showTitles',
                                        'title',
                                        'subtitle',
                                        'footnote'
                                    ];

                                    // Dynamically export all properties except excluded ones
                                    var settings = {};
                                    for (var key in data) {
                                        if (data.hasOwnProperty(key) && excludeProps.indexOf(key) === -1) {
                                            settings[key] = data[key];
                                        }
                                    }

                                    // Create export object
                                    var exportData = {
                                        exportVersion: "2.0",
                                        exportDate: new Date().toISOString(),
                                        extensionVersion: "2.0.58",
                                        settings: settings
                                    };

                                    // Convert to JSON string
                                    var jsonString = JSON.stringify(exportData, null, 2);

                                    // Create blob and download
                                    var blob = new Blob([jsonString], { type: 'application/json' });
                                    var url = URL.createObjectURL(blob);
                                    var a = document.createElement('a');
                                    a.href = url;
                                    a.download = 'dynamic-content-sections-settings-' + new Date().toISOString().split('T')[0] + '.json';
                                    document.body.appendChild(a);
                                    a.click();
                                    document.body.removeChild(a);
                                    URL.revokeObjectURL(url);

                                    // Count exported items
                                    var groupCount = settings.groups ? settings.groups.length : 0;
                                    var sectionCount = settings.sections ? settings.sections.length : 0;
                                    var colorMappingCount = settings.colorMappings ? settings.colorMappings.length : 0;
                                    var colorMapCount = settings.colorMapsList ? settings.colorMapsList.length : 0;
                                    var totalProps = Object.keys(settings).length;

                                    alert('✅ Settings exported successfully!\n\n' +
                                          'File: ' + a.download + '\n\n' +
                                          'Exported:\n' +
                                          '- Groups: ' + groupCount + '\n' +
                                          '- Sections: ' + sectionCount + '\n' +
                                          '- Color Mappings: ' + colorMappingCount + '\n' +
                                          '- Color Maps: ' + colorMapCount + '\n' +
                                          '- Total properties: ' + totalProps);
                                } catch (error) {
                                    alert('❌ Export failed: ' + error.message);
                                }
                            }
                        },
                        importButton: {
                            component: "button",
                            label: "Import Settings from File",
                            action: function(data) {
                                try {
                                    // Create file input
                                    var input = document.createElement('input');
                                    input.type = 'file';
                                    input.accept = '.json';

                                    input.onchange = function(e) {
                                        var file = e.target.files[0];
                                        if (!file) return;

                                        var reader = new FileReader();
                                        reader.onload = function(event) {
                                            try {
                                                var importData = JSON.parse(event.target.result);

                                                // Validate import data
                                                if (!importData.exportVersion || !importData.settings) {
                                                    throw new Error('Invalid file format');
                                                }

                                                // Get current model
                                                var app = qlik.currApp();
                                                var objectId = data.qInfo.qId;

                                                app.getObject(objectId).then(function(model) {
                                                    // Apply all settings dynamically
                                                    var properties = model.properties;
                                                    var settings = importData.settings;

                                                    // Dynamically import all settings from the file
                                                    for (var key in settings) {
                                                        if (settings.hasOwnProperty(key)) {
                                                            properties[key] = settings[key];
                                                        }
                                                    }

                                                    // Save the model
                                                    model.setProperties(properties).then(function() {
                                                        // Count imported items
                                                        var groupCount = settings.groups ? settings.groups.length : 0;
                                                        var sectionCount = settings.sections ? settings.sections.length : 0;
                                                        var colorMappingCount = settings.colorMappings ? settings.colorMappings.length : 0;
                                                        var colorMapCount = settings.colorMapsList ? settings.colorMapsList.length : 0;
                                                        var totalProps = Object.keys(settings).length;

                                                        alert('✅ Settings imported successfully!\n\n' +
                                                              'Imported:\n' +
                                                              '- Groups: ' + groupCount + '\n' +
                                                              '- Sections: ' + sectionCount + '\n' +
                                                              '- Color Mappings: ' + colorMappingCount + '\n' +
                                                              '- Color Maps: ' + colorMapCount + '\n' +
                                                              '- Total properties: ' + totalProps + '\n\n' +
                                                              'Export version: ' + (importData.exportVersion || '1.0') + '\n' +
                                                              'Extension version: ' + (importData.extensionVersion || 'Unknown'));
                                                    }).catch(function(error) {
                                                        alert('❌ Failed to apply settings: ' + error.message);
                                                    });
                                                });
                                            } catch (error) {
                                                alert('Import failed: ' + error.message + '\n\nPlease ensure the file is a valid export from this extension.');
                                            }
                                        };
                                        reader.readAsText(file);
                                    };

                                    input.click();
                                } catch (error) {
                                    alert('Import failed: ' + error.message);
                                }
                            }
                        },
                        importWarning: {
                            component: "text",
                            label: "⚠️ Warning: Importing will replace ALL current settings. Make sure to export first if you want to keep a backup!"
                        }
                    }
                },
                about: {
                    type: "items",
                    label: "About",
                    items: {
                        title: {
                            component: {
                        template: '<div style="text-align: center; padding: 10px 0;"><h1 style="color: #003B5C; font-size: 28px; font-weight: bold; margin: 0;">Dynamic Content Sections</h1></div>'
                            }
                        },
                        version: {
                            component: {
                        template: '<div style="text-align: center; padding: 5px 0;"><span style="color: #555; font-size: 14px;">Version 2.0.58</span></div>'
                            }
                        },
                        spacer1: {
                            component: {
                        template: '<div style="text-align: center; padding: 15px 0;"><hr style="border: none; border-top: 1px solid #e0e0e0; width: 80%; margin: 0 auto;"></div>'
                            }
                        },
                        qlikLogo: {
                            component: {
                        template: '<div style="text-align: center; padding: 15px 0;"><img src="' + QLIK_LOGO_URL + '" style="max-width: 250px; height: auto;" alt="Qlik Sales Enablement"></div>'
                            }
                        },
                        copyright: {
                            component: {
                        template: '<div style="text-align: center; padding: 10px 0;"><span style="color: #666; font-size: 13px;">© Technical Sales Enablement</span></div>'
                            }
                        },
                        author: {
                            component: {
                        template: '<div style="text-align: center; padding: 5px 0;"><a href="mailto:globalsalesenablement@qlik.com" style="color: #009845; text-decoration: none; font-size: 13px;">Created by: Jochem Zwienenberg</a></div>'
                            }
                        },
                        spacer5: {
                            component: "text",
                            label: " "
                        }
                    }
                }
            }
        },

        paint: function($element, layout) {
            var self = this;
            var instanceId = this.backendApi.model.id;

            // DEV VERSION WATERMARK: Add visible indicator for development versions
            if (IS_DEV_VERSION) {
                // Remove any existing watermarks first
                $element.find('.dev-watermark, .dev-watermark-diagonal').remove();

                // Add corner badge watermark
                var $watermark = $('<div class="dev-watermark"></div>');
                $watermark.css({
                    'position': 'absolute',
                    'top': '10px',
                    'right': '10px',
                    'background': 'rgba(255, 152, 0, 0.95)',
                    'color': 'white',
                    'padding': '8px 16px',
                    'border-radius': '4px',
                    'font-weight': 'bold',
                    'font-size': '12px',
                    'z-index': '9999',
                    'box-shadow': '0 2px 8px rgba(0,0,0,0.3)',
                    'pointer-events': 'none',
                    'letter-spacing': '0.5px',
                    'border': '2px solid rgba(255, 152, 0, 1)',
                    'text-transform': 'uppercase'
                });
                $watermark.html('🔧 DEV VERSION');
                $element.append($watermark);

                // Add diagonal watermark overlay for extra visibility
                var $diagonalWatermark = $('<div class="dev-watermark-diagonal"></div>');
                $diagonalWatermark.css({
                    'position': 'absolute',
                    'top': '50%',
                    'left': '50%',
                    'transform': 'translate(-50%, -50%) rotate(-45deg)',
                    'font-size': '48px',
                    'font-weight': 'bold',
                    'color': 'rgba(255, 152, 0, 0.15)',
                    'pointer-events': 'none',
                    'z-index': '1',
                    'white-space': 'nowrap',
                    'text-transform': 'uppercase',
                    'letter-spacing': '4px',
                    'user-select': 'none'
                });
                $diagonalWatermark.html('DEVELOPMENT VERSION');
                $element.append($diagonalWatermark);
            }

            // MIGRATION: Automatically migrate old sections to groups structure (only once)
            if (!migrationDone[instanceId] && (!layout.groups || layout.groups.length === 0) && layout.sections && layout.sections.length > 0) {
                migrationDone[instanceId] = true; // Mark as migrated immediately

                // Show loading spinner during migration
                $element.empty();
                $element.css({
                    'position': 'relative',
                    'overflow-y': 'auto',
                    'overflow-x': 'hidden',
                    'height': '100%',
                    'background': 'transparent',
                    'font-family': layout.fontFamily || "'Source Sans Pro', sans-serif",
                    'font-size': (layout.fontSize || '14') + 'px'
                });
                $element.html('<div class="jz-loading-spinner">' +
                    '<div class="jz-loading-spinner-icon"></div>' +
                    '<div class="jz-loading-spinner-text">Migrating configuration...</div>' +
                    '<div class="jz-loading-spinner-subtext">Updating to new group structure</div>' +
                    '</div>');

                // Create a default group with all existing sections
                layout.groups = [{
                    groupLabel: "Migrated Sections",
                    groupWidth: "full",
                        sections: layout.sections
                }];

                // Update the model with migrated data
                // This will trigger paint() again, so we return here to avoid double render
                this.backendApi.model.setProperties({
                    groups: layout.groups,
                    sections: [] // Clear old sections array
                }).then(function() {

                }).catch(function(err) {

                });

                // Return early - setProperties will trigger paint() again with clean data
                return;
            }

            // Increment global paint counter - used to detect stale modal opens
            globalPaintCounter++;

            // CRITICAL: Invalidate all pending PK check callbacks
            currentPKCheckToken++;

            // Close all old PK check objects
            pkCheckObjects.forEach(function(obj) {
                try {
                    if (obj && typeof obj.close === 'function') {
                        obj.close();
                    }
                } catch (e) {
                    // Ignore errors
                }
            });
            pkCheckObjects = [];

            var paintId = 'paint_' + Date.now();

            try {
                var self = this;
                var app = qlik.currApp();

                // Show loading spinner ONLY on the very first render of this instance
                // Don't show on subsequent renders (e.g., when opening properties panel)
                if (!hasRenderedOnce[instanceId]) {
                    $element.empty();
                    $element.css({
                        'position': 'relative',
                        'overflow-y': 'auto',
                        'overflow-x': 'hidden',
                        'height': '100%',
                        'background': 'transparent',
                        'font-family': layout.fontFamily || "'Source Sans Pro', sans-serif",
                        'font-size': (layout.fontSize || '14') + 'px'
                    });
                    $element.html('<div class="jz-loading-spinner">' +
                        '<div class="jz-loading-spinner-icon"></div>' +
                        '<div class="jz-loading-spinner-text">Initializing extension...</div>' +
                        '<div class="jz-loading-spinner-subtext">Preparing to load content</div>' +
                        '</div>');
                }

                // Store model reference for button actions
                currentModel = this.backendApi.model;

                // Fetch current user email (only happens once, subsequent calls are ignored if already fetched)
                // CRITICAL: This must happen BEFORE the permission blocking check below
                if (MULTI_USER_ENABLED === 1 && currentUserEmail === 'user@qlik.com') {
                    fetchCurrentUser(layout, self);
                }

                // === CRITICAL: Wait for permissions to load before first render ===
                // This prevents flickering when edit buttons appear after permissions load
                if (!userPermissionsLoaded && (MULTI_USER_ENABLED === 1 || CLAUDE_ENABLED === 1) && layout.mcpOrchestratorUrl) {
                    // Don't proceed with render - permissions will trigger repaint when loaded
                    return;
                }

                // Simple loop protection with much higher threshold
                // This allows property changes to go through while still preventing true infinite loops
                paintCallCount++;
                if (paintCallResetTimer) clearTimeout(paintCallResetTimer);
                paintCallResetTimer = setTimeout(function() {
                    paintCallCount = 0;
                }, 2000);  // Increased from 1s to 2s

                // Only block if we hit an extreme threshold (50+ calls in 2 seconds)
                if (paintCallCount > 50) {

                    return;
                }

                // CRITICAL: Never block paint during normal operations
                // Only block if save is in progress AND it's been less than 500ms since save started
                if (savingInProgress) {

                    return;
                }

                // REMOVED: Selection listener - Qlik calls paint() automatically on selection changes
                // The selection listener was causing flickering by triggering async operations
                // that interfered with the normal render cycle

            // ONE-TIME EVENT DELEGATION SETUP: Attach edit button handler ONCE per $element
            // This prevents multiple handlers from stacking up when renderSectionsUI is called multiple times
            if (!$element.data('edit-handler-attached')) {
                $element.on('click', '.section-edit-btn', function(e) {
                    e.stopPropagation();

                    var $btn = $(this);
                    var section = $btn.data('section-obj');
                    var sectionData = $btn.data('section-data-obj');
                    var $section = $btn.data('section-element');
                    var $content = $btn.data('content-element');

                    // Check if button is disabled
                    if ($btn.prop('disabled')) {
                        return;
                    }

                    // Capture paint counter at time of click
                    var clickPaintCounter = globalPaintCounter;

                    // Add DOM references so we can update the section after save
                    sectionData.$element = $section;
                    sectionData.$content = $content;

                    // Pass the currently rendered HTML content (what user sees) AND original template
                    var currentHtml = $content.html();
                    var originalTemplateHtml = $btn.data('original-template-html') || currentHtml;
                    openSectionEditModal(section, sectionData, layout, self, currentHtml, originalTemplateHtml, clickPaintCounter, $element, sectionsData, instanceId);
                });

                $element.data('edit-handler-attached', true);
            }

            // Check if we're already rendering this instance
            // BUT allow re-render if it's been more than 5 seconds (likely stuck)
            var now = Date.now();
            if (renderingState[instanceId]) {
                var lastRenderStart = renderStartTime[instanceId];
                if (lastRenderStart && (now - lastRenderStart) < 5000) {
                    return;
                }
                // Reset stuck render state
                renderingState[instanceId] = false;
                delete renderStartTime[instanceId];
            }

            // CRITICAL: Destroy all previous generic objects for this instance
            // This prevents orphaned callbacks from triggering stale renders
            if (genericObjectsRegistry[instanceId]) {
                genericObjectsRegistry[instanceId].forEach(function(obj) {
                    try {
                        if (obj && typeof obj.close === 'function') {
                            obj.close();
                        }
                    } catch (e) {
                        // Ignore errors during cleanup
                    }
                });
            }
            genericObjectsRegistry[instanceId] = [];

            // Mark as rendering and create new render token
            var now = Date.now();
            renderingState[instanceId] = true;
            renderStartTime[instanceId] = now;
            var renderToken = now + Math.random();
            currentRenderToken[instanceId] = renderToken;

            // Safety timeout: if rendering doesn't complete in 10 seconds, reset state
            setTimeout(function() {
                if (renderingState[instanceId] === true && currentRenderToken[instanceId] === renderToken) {
                    
                    renderingState[instanceId] = false;
                    delete renderStartTime[instanceId];
                }
            }, 10000);

            // IMMEDIATE: Close modal and disable all edit buttons when paint() starts
            // This prevents modal flickering when selections change
            var hadOpenModal = $('.edit-modal-overlay').length > 0;
            if (hadOpenModal) {
                closeModalAndCleanup();
            }

            // Disable all edit buttons in this element during render
            $element.find('.section-edit-btn').prop('disabled', true).addClass('disabled');

            // No PK cache checking here - we'll check PK value when needed in renderSectionsWithData

            // Handle backward compatibility: migrate old sections to groups structure
            var groups = layout.groups || [];
            if (groups.length === 0 && layout.sections && layout.sections.length > 0) {
                // Migrate old sections to a default group
                groups = [{
                    groupLabel: "Default Group",
                    groupWidth: "full",
                        sections: layout.sections
                }];
            }

            // Create per-section hypercubes by parsing master item references
            // Flatten all sections from all groups
            var allSections = [];
            var globalSpacing = layout.spacing !== undefined ? layout.spacing : 5;
            groups.forEach(function(group, groupIndex) {
                if (group.sections && group.sections.length > 0) {
                    group.sections.forEach(function(section) {
                        // Tag section with its parent group info
                        // Use groupIndex to ensure uniqueness even if labels are identical
                        section._groupLabel = group.groupLabel;
                        section._groupWidth = group.groupWidth;
                        section._groupBgColor = group.groupBgColor || 'transparent';
                        section._groupSpacing = group.groupSpacing !== undefined ? group.groupSpacing : globalSpacing;
                        section._groupBorderColor = group.groupBorderColor || 'transparent';
                        section._groupBorderWidth = group.groupBorderWidth || 0;
                        section._groupBorderStyle = group.groupBorderStyle || 'solid';
                        section._groupIndex = groupIndex;
                        allSections.push(section);
                    });
                }
            });

            var sections = allSections;
            var sectionsData = [];
            var sectionsLoaded = 0;

            if (sections.length === 0) {
                $element.html('<div style="padding: 20px; text-align: center; color: #999;">No groups or sections configured. Add groups in the properties panel.</div>');
                renderingState[instanceId] = false;
                delete renderStartTime[instanceId];
                return;
            }

            // Set timeout to render even if master items don't load
            var renderTimeout = setTimeout(function() {
                if (sectionsLoaded === 0 && sectionsData.length === 0) {
                    
                    // Force render with empty data
                    sections.forEach(function(section) {
                        sectionsData.push({
                            section: section,
                            data: null,
                            numDimensions: 0,
                            numMeasures: 0,
                            itemMapping: {},
                            notFoundItems: []  // Track that items weren't loaded
                        });
                    });
                    // Check if this is still the current render (not stale)
                    if (currentRenderToken[instanceId] === renderToken) {
                        renderSectionsWithData($element, layout, sectionsData, self, instanceId);
                    }
                }
            }, 5000);  // Increased from 2s to 5s

            // Check if we have cached master items
            if (masterItemsCache.cached && masterItemsCache.dimensions && masterItemsCache.measures) {
                // Use cached data - skip API calls
                clearTimeout(renderTimeout);

                var dimensionList = masterItemsCache.dimensions;
                var measureList = masterItemsCache.measures;

                // Build lookup maps from cached data
                var dimensionMap = {};
                var measureMap = {};

                if (dimensionList && dimensionList.length > 0) {
                    dimensionList.forEach(function(dim) {
                        dimensionMap[dim.qMeta.title] = dim.qInfo.qId;
                    });
                }

                if (measureList && measureList.length > 0) {
                    measureList.forEach(function(mea) {
                        measureMap[mea.qMeta.title] = mea.qInfo.qId;
                    });
                }

                updateLoadingMessage($element, 'Processing sections...', 'Using cached master items (' + dimensionList.length + ' dimensions, ' + measureList.length + ' measures)');

                // Process sections using cached data (same logic as fresh fetch)
                sections.forEach(function(section, idx) {
                    var sectionData = {
                        section: section,
                        data: null,
                        numDimensions: 0,
                        numMeasures: 0
                    };

                    var markdownText = section.markdownText || '';
                    var masterItemPattern = /\{\{(\[.*?\])\}\}/g;
                    var matches = [];
                    var match;

                    while ((match = masterItemPattern.exec(markdownText)) !== null) {
                        var itemName = match[1].slice(1, -1);
                        matches.push(itemName);
                    }

                    var colorByPattern = /colorBy="\[([^\]]+)\]"/g;
                    while ((match = colorByPattern.exec(markdownText)) !== null) {
                        var itemName = match[1];
                        matches.push(itemName);
                    }

                    var uniqueItems = [];
                    matches.forEach(function(item) {
                        if (uniqueItems.indexOf(item) === -1) {
                            uniqueItems.push(item);
                        }
                    });

                    var qDimensions = [];
                    var qMeasures = [];
                    var itemMapping = {};
                    var notFoundItems = [];
                    var dimIndex = 0;
                    var meaIndex = 0;

                    uniqueItems.forEach(function(itemName) {
                        if (dimensionMap[itemName]) {
                            qDimensions.push({ qLibraryId: dimensionMap[itemName] });
                            itemMapping[itemName] = { type: 'dim', index: dimIndex++ };
                        } else if (measureMap[itemName]) {
                            qMeasures.push({ qLibraryId: measureMap[itemName] });
                            itemMapping[itemName] = { type: 'mea', index: meaIndex++ };
                        } else {
                            notFoundItems.push(itemName);
                            // console.warn('[MASTER ITEM] NOT FOUND:', itemName);
                        }
                    });

                    sectionData.notFoundItems = notFoundItems;

                    if (qDimensions.length > 0 || qMeasures.length > 0) {
                        updateLoadingMessage($element, 'Fetching data from Qlik...', 'Loading section ' + (idx + 1) + ' of ' + sections.length);

                        var cubeDef = {
                            qDimensions: qDimensions,
                            qMeasures: qMeasures,
                            qInitialDataFetch: [{
                                qTop: 0,
                                qLeft: 0,
                                qHeight: 100,
                                qWidth: 20
                            }]
                        };

                        app.createCube(cubeDef, function(reply) {
                            if (currentRenderToken[instanceId] !== renderToken) {
                                return;
                            }

                            sectionData.data = reply.qHyperCube && reply.qHyperCube.qDataPages ? reply.qHyperCube.qDataPages[0] : null;
                            sectionData.numDimensions = reply.qHyperCube && reply.qHyperCube.qDimensionInfo ? reply.qHyperCube.qDimensionInfo.length : 0;
                            sectionData.numMeasures = reply.qHyperCube && reply.qHyperCube.qMeasureInfo ? reply.qHyperCube.qMeasureInfo.length : 0;
                            sectionData.itemMapping = itemMapping;

                            sectionsData[idx] = sectionData;
                            sectionsLoaded++;

                            if (sectionsLoaded === sections.length) {
                                if (currentRenderToken[instanceId] === renderToken) {
                                    renderSectionsWithData($element, layout, sectionsData, self, instanceId);
                                }
                            }
                        });
                    } else {
                        sectionsData[idx] = sectionData;
                        sectionsLoaded++;

                        if (sectionsLoaded === sections.length) {
                            if (currentRenderToken[instanceId] === renderToken) {
                                renderSectionsWithData($element, layout, sectionsData, self, instanceId);
                            }
                        }
                    }
                });
            } else {
                // First time or cache invalidated - fetch from Qlik
                updateLoadingMessage($element, 'Fetching data from Qlik...', 'Loading master items');

            var dimListDef = {
                qInfo: { qType: 'DimensionList' },
                qDimensionListDef: { qType: 'dimension', qData: { title: '/qMetaDef/title' } }
            };

            var measureListDef = {
                qInfo: { qType: 'MeasureList' },
                qMeasureListDef: { qType: 'measure', qData: { title: '/qMetaDef/title' } }
            };

            app.createGenericObject(dimListDef, function(dimensionReply) {
                // Track this object for cleanup
                if (dimensionReply && dimensionReply.model) {
                    genericObjectsRegistry[instanceId].push(dimensionReply.model);
                }

                // Check if this render is still current
                if (currentRenderToken[instanceId] !== renderToken) {
                    return;
                }

                app.createGenericObject(measureListDef, function(measureReply) {
                    // Track this object for cleanup
                    if (measureReply && measureReply.model) {
                        genericObjectsRegistry[instanceId].push(measureReply.model);
                    }

                    // Check if this render is still current
                    if (currentRenderToken[instanceId] !== renderToken) {
                        return;
                    }

                    clearTimeout(renderTimeout);

                    // Build lookup maps
                    var dimensionMap = {};
                    var measureMap = {};

                    // Extract dimensions from reply
                    var dimensionList = dimensionReply && dimensionReply.qDimensionList && dimensionReply.qDimensionList.qItems ? dimensionReply.qDimensionList.qItems : [];
                    var measureList = measureReply && measureReply.qMeasureList && measureReply.qMeasureList.qItems ? measureReply.qMeasureList.qItems : [];

                    if (dimensionList && dimensionList.length > 0) {
                        dimensionList.forEach(function(dim) {
                            dimensionMap[dim.qMeta.title] = dim.qInfo.qId;
                        });
                    }

                    if (measureList && measureList.length > 0) {
                        measureList.forEach(function(mea) {
                            measureMap[mea.qMeta.title] = mea.qInfo.qId;
                        });
                    }

                            // Cache master items to avoid re-fetching on every paint
                            masterItemsCache.dimensions = dimensionList;
                            masterItemsCache.measures = measureList;
                            masterItemsCache.cached = true;

                            // Update loading message
                            updateLoadingMessage($element, 'Processing sections...', 'Found ' + dimensionList.length + ' dimensions and ' + measureList.length + ' measures');

                            // Process each section
                    sections.forEach(function(section, idx) {
                        var sectionData = {
                            section: section,
                            data: null,
                            numDimensions: 0,
                            numMeasures: 0
                        };

                        // Extract master item references from markdown: {{[Master Item Name]}}
                        var markdownText = section.markdownText || '';

                        // LOG: Section markdown being processed
                        // 

                        var masterItemPattern = /\{\{(\[.*?\])\}\}/g;
                        var matches = [];
                        var match;

                        while ((match = masterItemPattern.exec(markdownText)) !== null) {
                            var itemName = match[1].slice(1, -1); // Remove [ and ]
                            matches.push(itemName);
                        }

                        // Also extract colorBy="[Master Item Name]" from content tags
                        var colorByPattern = /colorBy="\[([^\]]+)\]"/g;
                        while ((match = colorByPattern.exec(markdownText)) !== null) {
                            var itemName = match[1];
                            matches.push(itemName);
                        }

                        // LOG: Pattern matching results

                        // Remove duplicates
                        var uniqueItems = [];
                        matches.forEach(function(item) {
                            if (uniqueItems.indexOf(item) === -1) {
                        uniqueItems.push(item);
                            }
                        });

                        // Build hypercube definition and track master item names
                        var qDimensions = [];
                        var qMeasures = [];
                        var itemMapping = {}; // Maps master item name to {type: 'dim'/'mea', index: 0/1/2}
                        var notFoundItems = []; // Track items that weren't found

                        var dimIndex = 0;
                        var meaIndex = 0;

                        uniqueItems.forEach(function(itemName) {
                            if (dimensionMap[itemName]) {
                        qDimensions.push({ qLibraryId: dimensionMap[itemName] });
                        itemMapping[itemName] = { type: 'dim', index: dimIndex++ };
                            } else if (measureMap[itemName]) {
                        qMeasures.push({ qLibraryId: measureMap[itemName] });
                        itemMapping[itemName] = { type: 'mea', index: meaIndex++ };
                            } else {
                        notFoundItems.push(itemName);
                        // console.warn('[MASTER ITEM] NOT FOUND:', itemName);
                            }
                        });

                        // Store not found items for error display
                        sectionData.notFoundItems = notFoundItems;

                        // LOG: Summary before hypercube creation
                        // 
                        // 
                        // if (notFoundItems.length > 0) {
                        //     
                        // }

                        if (qDimensions.length > 0 || qMeasures.length > 0) {
                            // Update loading message when fetching section data
                            updateLoadingMessage($element, 'Fetching data from Qlik...', 'Loading section ' + (idx + 1) + ' of ' + sections.length);

                            // Create hypercube for this section
                            var cubeDef = {
                        qDimensions: qDimensions,
                        qMeasures: qMeasures,
                        qInitialDataFetch: [{
                                    qTop: 0,
                                    qLeft: 0,
                                    qHeight: 100,
                                    qWidth: 20
                                }]
                            };

                            // LOG: Cube definition before creation
                            //
                            //
                            //
                            if (notFoundItems.length > 0) {

                            }

                            app.createCube(cubeDef, function(reply) {
                                // CRITICAL: Check if this render is still current
                        if (currentRenderToken[instanceId] !== renderToken) {
                                    return;
                                }

                                // LOG: Hypercube response
                                // 
                        if (reply.qHyperCube) {
                                    // 
                                    // 
                                    // 
                                    // 
                                }

                        sectionData.data = reply.qHyperCube && reply.qHyperCube.qDataPages ? reply.qHyperCube.qDataPages[0] : null;
                        sectionData.numDimensions = reply.qHyperCube && reply.qHyperCube.qDimensionInfo ? reply.qHyperCube.qDimensionInfo.length : 0;
                        sectionData.numMeasures = reply.qHyperCube && reply.qHyperCube.qMeasureInfo ? reply.qHyperCube.qMeasureInfo.length : 0;
                        sectionData.itemMapping = itemMapping;

                                // 
                                // 

                        sectionsData[idx] = sectionData;
                                // 
                        sectionsLoaded++;

                        if (sectionsLoaded === sections.length) {
                                    // Check if this is still the current render (not stale)
                                    if (currentRenderToken[instanceId] === renderToken) {
                                        renderSectionsWithData($element, layout, sectionsData, self, instanceId);
                                    }
                                }
                            });
                        } else {
                            // No master items referenced
                            sectionsData[idx] = sectionData;
                            sectionsLoaded++;

                            if (sectionsLoaded === sections.length) {
                                // Check if this is still the current render (not stale)
                        if (currentRenderToken[instanceId] === renderToken) {
                                    renderSectionsWithData($element, layout, sectionsData, self, instanceId);
                                }
                            }
                        }
                    });
                });
            });
            }
        } catch (error) {
                
                
                $element.html('<div style="padding: 20px; text-align: center; color: #C62828; background: #FFEBEE; border-radius: 4px; margin: 10px;">' +
                    '<strong>⚠️ Extension Error</strong><br>' +
                    'Error: ' + error.message + '<br>' +
                    '<small style="color: #666; margin-top: 10px; display: block;">Check browser console for details</small>' +
                    '</div>');
                // Reset rendering state so future paints can try again
                if (this.backendApi && this.backendApi.model) {
                    var instanceId = this.backendApi.model.id;
                    renderingState[instanceId] = false;
                    delete renderStartTime[instanceId];
                }
            }
        }
    };

    // Parse markdown with color syntax
    // Helper function to convert Qlik color formats to valid CSS
    function normalizeColor(color) {
        if (!color) return color;

        // Handle RGB(r,g,b) or rgb(r,g,b) format from Qlik
        var rgbMatch = color.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/i);
        if (rgbMatch) {
            return 'rgb(' + rgbMatch[1] + ',' + rgbMatch[2] + ',' + rgbMatch[3] + ')';
        }

        // Handle ARGB(a,r,g,b) format from Qlik
        var argbMatch = color.match(/argb\((\d+),\s*(\d+),\s*(\d+),\s*(\d+)\)/i);
        if (argbMatch) {
            var alpha = parseInt(argbMatch[1]) / 255;
            return 'rgba(' + argbMatch[2] + ',' + argbMatch[3] + ',' + argbMatch[4] + ',' + alpha + ')';
        }

        // Map common color names to darker/better versions
        var colorMap = {
            'yellow': '#B8860B',    // Darker yellow (DarkGoldenrod)
            'orange': '#FF8C00',    // DarkOrange
            'lightgreen': '#32CD32', // LimeGreen
            'red': '#DC143C',       // Crimson
            'green': '#228B22',     // ForestGreen
            'blue': '#1E90FF',      // DodgerBlue
            'gray': '#696969'       // DimGray
        };

        var lowerColor = color.toLowerCase();
        if (colorMap[lowerColor]) {
            return colorMap[lowerColor];
        }

        // Already valid format (hex, color name)
        return color;
    }

    function parseMarkdown(text) {
        if (!text) return '';

        return text
            // Headers (process first)
            .replace(/^#### (.*$)/gim, '<h4>$1</h4>')
            .replace(/^### (.*$)/gim, '<h3>$1</h3>')
            .replace(/^## (.*$)/gim, '<h2>$1</h2>')
            .replace(/^# (.*$)/gim, '<h1>$1</h1>')
            // Bold and Italic (process before color syntax so **{color:text}** works)
            .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
            .replace(/\*(.*?)\*/g, '<em>$1</em>')
            // Badge syntax: {badge=color:text} - creates colored badge with background
            .replace(/\{badge=([^}:]+):(.*?)\}/g, function(match, color, content) {
                var colorName = color.trim().toLowerCase();
                var bgColor, textColor;

                // Predefined badge colors
                var badgeColors = {
                    'yellow': { bg: '#FFF3CD', text: '#856404' },
                    'orange': { bg: '#FFE5B4', text: '#8B4513' },
                    'red': { bg: '#FFCCCB', text: '#8B0000' },
                    'lightgreen': { bg: '#E8F5E9', text: '#2E7D32' },
                    'green': { bg: '#C8E6C9', text: '#1B5E20' },
                    'blue': { bg: '#D1ECF1', text: '#0C5460' },
                    'gray': { bg: '#E2E3E5', text: '#383D41' }
                };

                if (badgeColors[colorName]) {
                    bgColor = badgeColors[colorName].bg;
                    textColor = badgeColors[colorName].text;
                } else {
                    bgColor = normalizeColor(color.trim());
                    textColor = '#000';
                }

                return '<span style="background-color: ' + bgColor + '; color: ' + textColor + '; padding: 2px 8px; border-radius: 3px; font-size: 13px; display: inline-block; margin: 0 2px;">' + content + '</span>';
            })
            // Color syntax: {color:value} - now supports RGB(), ARGB(), hex, and color names
            .replace(/\{([^}:]+):(.*?)\}/g, function(match, color, content) {
                var normalizedColor = normalizeColor(color.trim());
                return '<span style="color: ' + normalizedColor + '">' + content + '</span>';
            })
            // Line breaks
            .replace(/\n\n/g, '<br><br>')
            .replace(/\n/g, '<br>');
    }

    // Process content tags in markdown
    function processContentTags(markdownText, data, numDimensions, numMeasures, layout, colorByDimension, itemMapping) {
        if (!markdownText) return '';

        colorByDimension = colorByDimension || null;
        itemMapping = itemMapping || {};

        // 
        // 
        // 

        // Process #[image] tags FIRST (before title tags so they can be embedded inline)
        markdownText = markdownText.replace(/#\[image\s*((?:[^\]"']|["'][^"']*["'])*)\]#\[\/image\]/gi, function(match, attrs) {
            var src = (attrs.match(/src=["']([^"']+)["']/i) || [])[1] || '';
            var alt = (attrs.match(/alt=["']([^"']+)["']/i) || [])[1] || '';
            var align = (attrs.match(/align=["']([^"']+)["']/i) || [])[1] || 'left';
            var width = (attrs.match(/width=["']([^"']+)["']/i) || [])[1] || 'auto';
            var height = (attrs.match(/height=["']([^"']+)["']/i) || [])[1] || 'auto';
            var maxWidth = (attrs.match(/maxWidth=["']([^"']+)["']/i) || [])[1] || '';
            var maxHeight = (attrs.match(/maxHeight=["']([^"']+)["']/i) || [])[1] || '';
            var keepRatio = (attrs.match(/keepRatio=["']?(true|false)["']?/i) || [])[1] || 'false';
            return processImageTag(src, alt, align, width, height, maxWidth, maxHeight, keepRatio === 'true', layout);
        });

        // Process {bg:...}...{/bg} tags (universal formatting containers)
        // IMPORTANT: Use [\s\S]*? to match content across newlines
        markdownText = markdownText.replace(/\{bg:([^}]+)\}([\s\S]*?)\{\/bg\}/g, function(match, attrs, content) {
            return processBackgroundContainerTag(attrs, content);
        });

        // Process #[title] tags (consume surrounding whitespace to prevent extra spacing)
        // IMPORTANT: Use [\s\S]*? to match content across newlines
        markdownText = markdownText.replace(/\s*#\[title\s*((?:[^\]"']|["'][^"']*["'])*)\]([\s\S]*?)#\[\/title\]\s*/g, function(match, attrs, content) {
            var separator = (attrs.match(/separator=["']([^"']+)["']/i) || [])[1] === 'true';
            var font = (attrs.match(/font=["']([^"']+)["']/i) || [])[1];
            var size = (attrs.match(/size=["']([^"']+)["']/i) || [])[1];
            var style = (attrs.match(/style=["']([^"']+)["']/i) || [])[1];
            var align = (attrs.match(/align=["']([^"']+)["']/i) || [])[1] || 'left';
            return processTitleTag(content, separator, font, size, style, align);
        });

        // Process #[colortext] tags (apply color maps to values)
        markdownText = markdownText.replace(/#\[colortext\s*((?:[^\]"]|"[^"]*")*)\]([\s\S]*?)#\[\/colortext\]/g, function(match, attrs, content) {
            var label = (attrs.match(/label="([^"]+)"/) || [])[1] || '';
            var map = (attrs.match(/map="([^"]+)"/) || [])[1] || '';
            var mapDef = (attrs.match(/mapDef="([^"]+)"/) || [])[1] || ''; // Inline map definition
            var extractPattern = (attrs.match(/extract="([^"]+)"/) || [])[1] || '\\(([a-zA-Z0-9]+)\\)';
            return processColorTextTag(content, label, map, mapDef, extractPattern, layout);
        });

        // Process #[table] tags
        markdownText = markdownText.replace(/#\[table\s*((?:[^\]"]|"[^"]*")*)\]([\s\S]*?)#\[\/table\]/g, function(match, attrs, content) {
            var headerColor = (attrs.match(/headerColor="([^"]+)"/) || [])[1];
            var stripedRows = (attrs.match(/stripedRows="([^"]+)"/) || [])[1];
            var hideHeader = (attrs.match(/hideHeader="([^"]+)"/) || [])[1] === 'true';
            var font = (attrs.match(/font="([^"]+)"/) || [])[1];
            var size = (attrs.match(/size="([^"]+)"/) || [])[1];
            var style = (attrs.match(/style="([^"]+)"/) || [])[1];
            var align = (attrs.match(/align="([^"]+)"/) || [])[1] || 'left';
            var sort = (attrs.match(/sort="([^"]+)"/) || [])[1];
            var widths = (attrs.match(/widths="([^"]+)"/) || [])[1]; // Manual column widths (comma-separated percentages)
            return processTableTag(content, data, numDimensions, numMeasures, layout, headerColor, stripedRows, hideHeader, font, size, style, align, sort, widths, itemMapping);
        });

        // Process #[list] tags
        // Pattern explanation: Match attributes that can contain quoted strings with brackets inside
        // IMPORTANT: Use [\s\S]*? to match content across newlines
        markdownText = markdownText.replace(/#\[list\s*((?:[^\]"]|"[^"]*")*)\]([\s\S]*?)#\[\/list\]/g, function(match, attrs, content) {
            var type = (attrs.match(/type="([^"]+)"/) || [])[1] || 'bulleted';
            var columns = parseInt((attrs.match(/columns="([^"]+)"/) || [])[1] || '1', 10);
            var numbering = (attrs.match(/numbering="([^"]+)"/) || [])[1] || 'horizontal';
            var dividers = (attrs.match(/dividers="([^"]+)"/) || [])[1] === 'true';
            var font = (attrs.match(/font="([^"]+)"/) || [])[1];
            var size = (attrs.match(/size="([^"]+)"/) || [])[1];
            var style = (attrs.match(/style="([^"]+)"/) || [])[1];
            var align = (attrs.match(/align="([^"]+)"/) || [])[1] || 'left';
            var sort = (attrs.match(/sort="([^"]+)"/) || [])[1];

            // Extract colorBy attribute: colorBy="[Master Item Name]"
            var colorByMatch = attrs.match(/colorBy="\[([^\]]+)\]"/);
            var tagColorBy = colorByMatch ? colorByMatch[1] : null;

            var result = processListTag(content, data, numDimensions, numMeasures, layout, type, columns, numbering, dividers, font, size, style, align, sort, tagColorBy, itemMapping);
            return result;
        });

        // Process #[concat] tags
        // IMPORTANT: Use [\s\S]*? to match content across newlines
        markdownText = markdownText.replace(/#\[concat\s*((?:[^\]"]|"[^"]*")*)\]([\s\S]*?)#\[\/concat\]/g, function(match, attrs, content) {
            var delimiter = (attrs.match(/delimiter="([^"]+)"/) || [])[1] || ', ';
            var font = (attrs.match(/font="([^"]+)"/) || [])[1];
            var size = (attrs.match(/size="([^"]+)"/) || [])[1];
            var style = (attrs.match(/style="([^"]+)"/) || [])[1];
            return processConcatTag(content, data, numDimensions, numMeasures, layout, delimiter, font, size, style, itemMapping);
        });

        // Process #[grid] tags
        // IMPORTANT: Use [\s\S]*? to match content across newlines
        markdownText = markdownText.replace(/#\[grid\s*((?:[^\]"]|"[^"]*")*)\]([\s\S]*?)#\[\/grid\]/g, function(match, attrs, content) {
            var columns = parseInt((attrs.match(/columns="([^"]+)"/) || [])[1] || '2', 10);
            var dividers = (attrs.match(/dividers="([^"]+)"/) || [])[1] === 'true';
            var font = (attrs.match(/font="([^"]+)"/) || [])[1];
            var size = (attrs.match(/size="([^"]+)"/) || [])[1];
            var style = (attrs.match(/style="([^"]+)"/) || [])[1];
            return processGridTag(content, data, numDimensions, numMeasures, layout, columns, dividers, font, size, style, itemMapping);
        });

        // Process #[pivot] tags
        // IMPORTANT: Use [\s\S]*? to match content across newlines
        markdownText = markdownText.replace(/#\[pivot\s*((?:[^\]"]|"[^"]*")*)\]([\s\S]*?)#\[\/pivot\]/g, function(match, attrs, content) {
            var headerColor = (attrs.match(/headerColor="([^"]+)"/) || [])[1] || '#009845';
            var stripedRows = (attrs.match(/stripedRows="([^"]+)"/) || [])[1] === 'true';
            var font = (attrs.match(/font="([^"]+)"/) || [])[1];
            var size = (attrs.match(/size="([^"]+)"/) || [])[1];
            var style = (attrs.match(/style="([^"]+)"/) || [])[1];
            var align = (attrs.match(/align="([^"]+)"/) || [])[1] || 'left,right';
            return processPivotTag(content, data, numDimensions, numMeasures, layout, headerColor, stripedRows, font, size, style, align, itemMapping);
        });

        // Process #[kpi] tags
        // IMPORTANT: Use [\s\S]*? to match content across newlines
        markdownText = markdownText.replace(/#\[kpi\s*((?:[^\]"]|"[^"]*")*)\]([\s\S]*?)#\[\/kpi\]/g, function(match, attrs, content) {
            var label = (attrs.match(/label="([^"]+)"/) || [])[1] || '';
            var font = (attrs.match(/font="([^"]+)"/) || [])[1];
            var size = (attrs.match(/size="([^"]+)"/) || [])[1];
            var style = (attrs.match(/style="([^"]+)"/) || [])[1];
            return processKpiTag(content, data, numDimensions, numMeasures, layout, label, font, size, style, itemMapping);
        });

        // Process #[claude] or #[ai] tags for AI analysis (BEFORE #[box] so they can be nested)
        markdownText = markdownText.replace(/#\[(claude|ai)\s*((?:[^\]"]|"[^"]*")*)\]([\s\S]*?)#\[\/(claude|ai)\]/gi, function(match, tagName, attrs, content) {
            // Check permission instead of enableClaude checkbox
            if (!userHasFeaturePermission('jz_claude')) {
                return '<div style="padding: 8px; background: #fff3cd; border-left: 4px solid #ffc107; border-radius: 4px;">' +
                       '<strong>⚠️ Claude AI Access Denied</strong><br>You do not have permission to use Claude AI features.' +
                       '</div>';
            }
            var prompt = (attrs.match(/prompt="([^"]+)"/) || [])[1] || content.trim() || 'Please analyze this data.';
            var dataRefs = (attrs.match(/data="([^"]+)"/) || [])[1] || '';
            return processClaudeTag(prompt, dataRefs, content, data, numDimensions, numMeasures, layout);
        });

        // Process #[row] tags (for horizontal distribution of content)
        markdownText = markdownText.replace(/#\[row\s*((?:[^\]"]|"[^"]*")*)\]([\s\S]*?)#\[\/row\]/gi, function(match, attrs, content) {
            var align = (attrs.match(/align="([^"]+)"/i) || [])[1] || 'space-between';
            var gap = (attrs.match(/gap="([^"]+)"/i) || [])[1] || '15px';
            var center = (attrs.match(/center="([^"]+)"/i) || [])[1] === 'true';
            var font = (attrs.match(/font="([^"]+)"/) || [])[1];
            var size = (attrs.match(/size="([^"]+)"/) || [])[1];
            var style = (attrs.match(/style="([^"]+)"/) || [])[1];
            return processRowTag(content, align, gap, center, font, size, style);
        });

        // Process #[box] tags
        markdownText = markdownText.replace(/#\[box\s*((?:[^\]"]|"[^"]*")*)\]([\s\S]*?)#\[\/box\]/gi, function(match, attrs, content) {
            var color = (attrs.match(/color="([^"]+)"/i) || [])[1] || '#009845';
            var bgColor = (attrs.match(/bg[cC]olor="([^"]+)"/i) || [])[1] || '';
            var font = (attrs.match(/font="([^"]+)"/) || [])[1];
            var size = (attrs.match(/size="([^"]+)"/) || [])[1];
            var style = (attrs.match(/style="([^"]+)"/) || [])[1];
            var align = (attrs.match(/align="([^"]+)"/i) || [])[1] || 'left';
            var result = processBoxTag(content, color, bgColor, data, numDimensions, numMeasures, layout, font, size, style, align);
            return result;
        });

        // Process #[viz] tags (placeholders that will be replaced with actual visualizations)
        markdownText = markdownText.replace(/#\[viz\s*((?:[^\]"]|"[^"]*")*)\]#\[\/viz\]/g, function(match, attrs) {
            var vizId = (attrs.match(/id="([^"]+)"/) || [])[1];
            var vizName = (attrs.match(/name="([^"]+)"/) || [])[1];
            var height = (attrs.match(/height="([^"]+)"/) || [])[1] || '400px';
            return processVizTag(vizId, vizName, height);
        });

        return markdownText;
    }

    // Helper function to apply font, size, and style attributes
    function applyFontStyle(baseStyle, font, size, style) {
        var styleString = baseStyle || '';

        if (font) {
            styleString += ' font-family: ' + font + ';';
        }

        if (size) {
            styleString += ' font-size: ' + size + ';';
        }

        if (style) {
            var styles = style.split(',');
            styles.forEach(function(s) {
                s = s.trim().toLowerCase();
                if (s === 'bold') {
                    styleString += ' font-weight: bold;';
                } else if (s === 'italic') {
                    styleString += ' font-style: italic;';
                } else if (s === 'underline') {
                    styleString += ' text-decoration: underline;';
                }
            });
        }

        return styleString;
    }

    // Wrapper function to apply font styling to content
    function wrapWithFontStyle(html, font, size, style) {
        if (!font && !size && !style) {
            return html; // No styling needed
        }

        var wrapperStyle = applyFontStyle('width: 100%; display: block; box-sizing: border-box;', font, size, style);
        return '<div style="' + wrapperStyle + '">' + html + '</div>';
    }

    function processBackgroundContainerTag(attrs, content) {
        // Parse attributes: bg:#006580 padding:10px align:center color:#fff
        var bgColor = (attrs.match(/^#[0-9a-fA-F]{6}|^#[0-9a-fA-F]{3}/) || [])[0] || 'transparent';
        var padding = (attrs.match(/padding:([^\s]+)/) || [])[1] || '10px';
        var align = (attrs.match(/align:([^\s]+)/) || [])[1] || 'left';
        var color = (attrs.match(/color:(#[0-9a-fA-F]{6}|#[0-9a-fA-F]{3}|[a-z]+)/) || [])[1] || '';
        var borderRadius = (attrs.match(/radius:([^\s]+)/) || [])[1] || '4px';
        var margin = (attrs.match(/margin:([^\s]+)/) || [])[1] || '0';

        // Convert \n to <br> for line breaks
        content = content.replace(/\n/g, '<br>');

        // Parse inline color/size/style syntax: {color:size:style:text}
        content = content.replace(/\{([^}:]+):([^}]*)\}/g, function(match, color, rest) {
            var normalizedColor = normalizeColor(color.trim());
            var parts = rest.split(':');
            var textContent, segmentSize, segmentStyle;

            if (parts.length >= 3) {
                var firstPart = parts[0].trim();
                var secondPart = parts[1].trim();
                var isSizeValue = firstPart && /\d+(px|pt|em|rem|%)/i.test(firstPart);
                var firstIsStyle = firstPart && /^(bold|italic|underline|bold,italic|italic,bold)$/i.test(firstPart);
                var secondIsStyle = secondPart && /^(bold|italic|underline|bold,italic|italic,bold)$/i.test(secondPart);

                if (isSizeValue && secondIsStyle) {
                    segmentSize = firstPart;
                    segmentStyle = secondPart;
                    textContent = parts.slice(2).join(':');
                } else if (firstIsStyle) {
                    segmentSize = null;
                    segmentStyle = firstPart;
                    textContent = parts.slice(1).join(':');
                } else if (isSizeValue) {
                    segmentSize = firstPart;
                    segmentStyle = null;
                    textContent = parts.slice(1).join(':');
                } else {
                    textContent = rest;
                    segmentSize = null;
                    segmentStyle = null;
                }
            } else if (parts.length === 2) {
                var firstPart = parts[0].trim();
                var isSizeValue = firstPart && /\d+(px|pt|em|rem|%)/i.test(firstPart);
                var isStyleValue = firstPart && /^(bold|italic|underline|bold,italic|italic,bold)$/i.test(firstPart);

                if (isSizeValue) {
                    segmentSize = firstPart;
                    segmentStyle = null;
                    textContent = parts[1];
                } else if (isStyleValue) {
                    segmentSize = null;
                    segmentStyle = firstPart;
                    textContent = parts[1];
                } else {
                    textContent = rest;
                    segmentSize = null;
                    segmentStyle = null;
                }
            } else {
                textContent = rest;
                segmentSize = null;
                segmentStyle = null;
            }

            var spanStyle = 'color: ' + normalizedColor + ';';
            if (segmentSize) spanStyle += ' font-size: ' + segmentSize + ';';
            if (segmentStyle) {
                var styles = segmentStyle.split(',');
                styles.forEach(function(s) {
                    s = s.trim().toLowerCase();
                    if (s === 'bold') spanStyle += ' font-weight: bold;';
                    else if (s === 'italic') spanStyle += ' font-style: italic;';
                    else if (s === 'underline') spanStyle += ' text-decoration: underline;';
                });
            }

            return '<span style="' + spanStyle + '">' + textContent + '</span>';
        });

        // Build style
        var style = 'background-color: ' + normalizeColor(bgColor) + '; ';
        style += 'padding: ' + padding + '; ';
        style += 'text-align: ' + align + '; ';
        style += 'margin: ' + margin + '; ';
        style += 'border-radius: ' + borderRadius + '; ';
        style += 'display: block; ';
        if (color) {
            style += 'color: ' + normalizeColor(color) + '; ';
        }

        return '<div style="' + style + '">' + content + '</div>';
    }

    function processTitleTag(content, separator, font, size, style, align) {
        // Split content by newlines to support multiple lines/paragraphs
        // Keep empty lines for spacing - trim individual lines but preserve empty lines
        var lines = content.split('\n').map(function(line) { return line.trim(); });

        // Trim only trailing empty lines (keep leading ones for top spacing)
        while (lines.length > 0 && !lines[lines.length - 1]) {
            lines.pop();
        }

        var html = '';

        lines.forEach(function(line, lineIndex) {
            // If line is empty, add spacing
            if (!line) {
                html += '<div style="height: 8px;"></div>';
                return;
            }
            // Helper function to parse color syntax
            function parseColorSyntax(text) {
                return text.replace(/\{([^}:]+):([^}]*)\}/g, function(match, color, rest) {
                    var normalizedColor = normalizeColor(color.trim());

                    // Check if rest contains additional colons (extended format)
                    var parts = rest.split(':');
                    var textContent, segmentSize, segmentStyle;

                    if (parts.length >= 3) {
                        // Could be {color:size:style:text}, {color::style:text}, or {color:style:text:with:colons}
                        var firstPart = parts[0].trim();
                        var secondPart = parts[1].trim();
                        var isSizeValue = firstPart && /\d+(px|pt|em|rem|%)/i.test(firstPart);
                        var firstIsStyle = firstPart && /^(bold|italic|underline|bold,italic|italic,bold)$/i.test(firstPart);
                        var secondIsStyle = secondPart && /^(bold|italic|underline|bold,italic|italic,bold)$/i.test(secondPart);

                        if (isSizeValue && secondIsStyle) {
                            // Full format: {color:size:style:text}
                            segmentSize = firstPart;
                            segmentStyle = secondPart;
                            textContent = parts.slice(2).join(':');
                        } else if (!firstPart && secondIsStyle) {
                            // Format: {color::style:text} (empty size)
                            segmentSize = null;
                            segmentStyle = secondPart;
                            textContent = parts.slice(2).join(':');
                        } else if (firstIsStyle) {
                            // Format: {color:style:text:with:colons} - text contains colons
                            segmentSize = null;
                            segmentStyle = firstPart;
                            textContent = parts.slice(1).join(':');
                        } else if (isSizeValue) {
                            // Format: {color:size:text:with:colons} - text contains colons
                            segmentSize = firstPart;
                            segmentStyle = null;
                            textContent = parts.slice(1).join(':');
                        } else {
                            // Simple format: {color:text:with:colons} - entire rest is text
                            textContent = rest;
                            segmentSize = null;
                            segmentStyle = null;
                        }
                    } else if (parts.length === 2) {
                        // Could be {color:size:text}, {color:style:text}, or {color:text}
                        var firstPart = parts[0].trim();
                        var isSizeValue = firstPart && /\d+(px|pt|em|rem|%)/i.test(firstPart);
                        var isStyleValue = firstPart && /^(bold|italic|underline|bold,italic|italic,bold)$/i.test(firstPart);

                        if (isSizeValue) {
                            // Format: {color:size:text}
                            segmentSize = firstPart;
                            segmentStyle = null;
                            textContent = parts[1];
                        } else if (isStyleValue) {
                            // Format: {color:style:text}
                            segmentSize = null;
                            segmentStyle = firstPart;
                            textContent = parts[1];
                        } else {
                            // Simple format: {color:text} (treat entire rest as text)
                            textContent = rest;
                            segmentSize = null;
                            segmentStyle = null;
                        }
                    } else {
                        // Simple format: {color:text} (no additional colons)
                        textContent = rest;
                        segmentSize = null;
                        segmentStyle = null;
                    }

                    // Build inline style for this segment
                    var segmentInlineStyle = 'color: ' + normalizedColor + ';';

                    if (segmentSize) {
                        segmentInlineStyle += ' font-size: ' + segmentSize + ';';
                    }

                    if (segmentStyle) {
                        var styles = segmentStyle.split(',').map(function(s) { return s.trim().toLowerCase(); });
                        styles.forEach(function(s) {
                            if (s === 'bold') {
                                segmentInlineStyle += ' font-weight: bold;';
                            } else if (s === 'italic') {
                                segmentInlineStyle += ' font-style: italic;';
                            } else if (s === 'underline') {
                                segmentInlineStyle += ' text-decoration: underline;';
                            }
                        });
                    }

                    return '<span style="' + segmentInlineStyle + '">' + textContent + '</span>';
                });
            }

            // Check if line contains split delimiter ||
            var hasSplit = line.indexOf('||') !== -1;

            if (hasSplit) {
                // Split layout: left side || right side
                var splitParts = line.split('||');
                var leftContent = parseColorSyntax(splitParts[0].trim());
                var rightContent = parseColorSyntax(splitParts[1].trim());

                // Build flexbox container with space-between
                var splitStyle = 'margin: 0; line-height: 1.3; display: flex; justify-content: space-between; align-items: center;';

                // Add margin-top for lines after the first (but not if previous line was empty - spacing already added)
                if (lineIndex > 0 && lines[lineIndex - 1]) {
                    splitStyle += ' margin-top: 4px;';
                }

                // Set default font size only if not provided at tag level
                if (!size) {
                    splitStyle = 'font-size: 18px; ' + splitStyle;
                }

                // Set default font weight only if not provided at tag level
                if (!style || style.indexOf('bold') === -1) {
                    splitStyle = 'font-weight: 600; ' + splitStyle;
                }

                splitStyle = applyFontStyle(splitStyle, font, size, style);

                html += '<div style="' + splitStyle + '">';
                html += '<div style="text-align: left;">' + leftContent + '</div>';
                html += '<div style="text-align: right;">' + rightContent + '</div>';
                html += '</div>';
            } else {
                // Normal single-column layout
                var titleText = parseColorSyntax(line);

                // Check if content contains inline images (inline-right or inline-left)
                var hasInlineImage = titleText.match(/inline-right|inline-left|display: inline-flex/);

                // Build title HTML with default styles
                var titleStyle = 'margin: 0; line-height: 1.3; text-align: ' + (align || 'left') + ';';

                // Add margin-top for lines after the first (but not if previous line was empty - spacing already added)
                if (lineIndex > 0 && lines[lineIndex - 1]) {
                    titleStyle += ' margin-top: 4px;';
                }

                // Set default font size only if not provided at tag level
                if (!size) {
                    titleStyle = 'font-size: 18px; ' + titleStyle;
                }

                // Set default font weight only if not provided at tag level
                if (!style || style.indexOf('bold') === -1) {
                    titleStyle = 'font-weight: 600; ' + titleStyle;
                }

                titleStyle = applyFontStyle(titleStyle, font, size, style);

                // If inline images are present, use flexbox layout to keep them on the same line
                if (hasInlineImage) {
                    titleStyle += ' display: flex; align-items: center; flex-wrap: nowrap;';
                    // Split content at the image div and wrap text separately
                    var parts = titleText.split(/(<div style="display: inline-flex[^>]*>.*?<\/div>)/i);
                    var wrappedText = '<span style="flex: 0 0 auto;">' + (parts[0] || '') + '</span>' + (parts[1] || '');
                    html += '<div style="' + titleStyle + '">' + wrappedText + '</div>';
                } else {
                    html += '<div style="' + titleStyle + '">' + titleText + '</div>';
                }
            }
        });

        // Add separator line if requested
        if (separator) {
            var separatorStyle = 'border-top: 1px solid #e0e0e0; margin: 2px 0 0 0;';
            html += '<div style="' + separatorStyle + '"></div>';
        }

        return html;
    }

    // Helper function to parse color map definition string
    // Format: "1:lightred,2:#FFC207,3:#FDFE17,4:#92D051,5:#25A661"
    // Extended format with background: "1:textColor:bgColor,2:textColor:bgColor"
    function parseColorMap(mapDefString) {
        var colorMap = {};
        if (!mapDefString) return colorMap;

        var pairs = mapDefString.split(',');
        pairs.forEach(function(pair) {
            var parts = pair.split(':');
            if (parts.length === 2) {
                // Old format: Key:TextColor
                var key = parts[0].trim();
                var color = parts[1].trim();
                colorMap[key] = { textColor: color, bgColor: null };
            } else if (parts.length === 3) {
                // New format: Key:TextColor:BgColor
                var key = parts[0].trim();
                var textColor = parts[1].trim();
                var bgColor = parts[2].trim();
                colorMap[key] = { textColor: textColor, bgColor: bgColor };
            }
        });

        return colorMap;
    }

    // Helper function to get color map from layout by name
    function getColorMapByName(mapName, layout) {
        if (!layout.colorMapsList || !mapName) return null;

        for (var i = 0; i < layout.colorMapsList.length; i++) {
            if (layout.colorMapsList[i].mapName === mapName) {
                return parseColorMap(layout.colorMapsList[i].mapDefinition);
            }
        }

        return null;
    }

    // Helper function to apply color map to a value
    function applyColorMap(value, colorMap, extractPattern) {
        if (!value || !colorMap) return value;

        // Extract the key from the value using the pattern
        var key = value;
        if (extractPattern) {
            try {
                var regex = new RegExp(extractPattern);
                var match = value.match(regex);
                if (match && match[1]) {
                    key = match[1];
                }
            } catch (e) {
                
            }
        }

        // Look up color
        var colorInfo = colorMap[key];
        if (colorInfo) {
            var styleStr = '';

            // Handle both old format (string) and new format (object)
            if (typeof colorInfo === 'string') {
                // Old format: just a color string
                styleStr = 'color: ' + normalizeColor(colorInfo);
            } else if (typeof colorInfo === 'object') {
                // New format: {textColor, bgColor}
                if (colorInfo.textColor) {
                    styleStr += 'color: ' + normalizeColor(colorInfo.textColor) + ';';
                }
                if (colorInfo.bgColor) {
                    styleStr += ' background-color: ' + normalizeColor(colorInfo.bgColor) + '; padding: 2px 6px; border-radius: 3px;';
                }
            }

            if (styleStr) {
                return '<span style="' + styleStr + '">' + value + '</span>';
            }
        }

        return value;
    }

    function processColorTextTag(content, label, mapName, mapDef, extractPattern, layout) {
        var value = content.trim();

        // Get color map (either by name from global maps, or inline definition)
        var colorMap = null;
        if (mapName) {
            colorMap = getColorMapByName(mapName, layout);
        } else if (mapDef) {
            colorMap = parseColorMap(mapDef);
        }

        if (!colorMap) {
            // No color map found, just return the value with optional label
            return label ? '<strong>' + label + ':</strong> ' + value : value;
        }

        // Apply color map
        var coloredValue = applyColorMap(value, colorMap, extractPattern);

        // Add label if provided
        if (label) {
            return '<strong>' + label + ':</strong> ' + coloredValue;
        }

        return coloredValue;
    }

    function processRowTag(content, align, gap, center, font, size, style) {
        // Split content by | to get individual cells
        var cells = content.trim().split('|').map(function(cell) { return cell.trim(); });

        // Map align values
        var justifyContent = 'space-between';
        if (align === 'left' || align === 'flex-start') justifyContent = 'flex-start';
        else if (align === 'center') justifyContent = 'center';
        else if (align === 'right' || align === 'flex-end') justifyContent = 'flex-end';
        else if (align === 'space-around') justifyContent = 'space-around';
        else if (align === 'space-evenly') justifyContent = 'space-evenly';

        var containerStyle = 'display: flex; justify-content: ' + justifyContent + '; gap: ' + gap + '; align-items: ' + (center ? 'center' : 'flex-start') + '; padding: 3px 0;';

        // Apply font styling to container
        if (font) containerStyle += ' font-family: ' + font + ';';
        if (size) containerStyle += ' font-size: ' + size + ';';
        if (style) {
            var styles = style.split(',');
            styles.forEach(function(s) {
                s = s.trim().toLowerCase();
                if (s === 'bold') containerStyle += ' font-weight: bold;';
                else if (s === 'italic') containerStyle += ' font-style: italic;';
                else if (s === 'underline') containerStyle += ' text-decoration: underline;';
            });
        }

        var html = '<div style="' + containerStyle + '">';

        // If aligning left/right/center, don't use flex:1 (which spreads items)
        var useFlex = (align === 'space-between' || align === 'space-around' || align === 'space-evenly');

        cells.forEach(function(cell) {
            var cellStyle = 'text-align: ' + (center ? 'center' : 'left') + ';';
            if (useFlex) {
                cellStyle += ' flex: 1;';
            }
            // Parse markdown (including color syntax) in each cell
            var parsedCell = parseMarkdown(cell);
            html += '<div style="' + cellStyle + '">' + parsedCell + '</div>';
        });
        html += '</div>';

        return html;
    }

    function processImageTag(src, alt, align, width, height, maxWidth, maxHeight, keepRatio, layout) {
        // Handle media library images (preferred method)
        // Syntax: media:filename.png or just filename.png (if not a full path)
        // Get actual app ID (not object ID)
        var appId = null;
        try {
            var app = qlik.currApp();
            appId = app ? app.id : null;
        } catch (e) {
            // console.error('Failed to get app ID:', e);
        }

        if (src.startsWith('media:')) {
            // Explicit media library reference: media:filename.png
            var filename = src.substring(6); // Remove 'media:' prefix
            // Use Qlik Cloud media API endpoint
            var tenantUrl = window.location.origin;
            if (appId) {
                src = tenantUrl + '/api/v1/apps/' + appId + '/media/files/' + filename;
            } else {
                // Fallback to old path if app ID unavailable (shouldn't happen)
                src = tenantUrl + '/content/default/' + filename;
            }
        } else if (!src.startsWith('http://') && !src.startsWith('https://') &&
                   !src.startsWith('/') && !src.startsWith('./') && !src.startsWith('../') &&
                   !src.startsWith('Img/') && !src.startsWith('data:')) {
            // If it's just a filename (no protocol/path), assume media library
            var tenantUrl = window.location.origin;
            if (appId) {
                src = tenantUrl + '/api/v1/apps/' + appId + '/media/files/' + src;
            } else {
                // Fallback to old path if app ID unavailable (shouldn't happen)
                src = tenantUrl + '/content/default/' + src;
            }
        }
        // Handle extension-relative paths
        else if (src.startsWith('./') || src.startsWith('../')) {
            src = '../../extensions/jz-dynamic-content-sections/' + src.replace(/^\.\.?\/?/, '');
        } else if (src.startsWith('Img/')) {
            src = '../../extensions/jz-dynamic-content-sections/' + src;
        }

        // Build style string
        var imgStyle = '';

        if (width && width !== 'auto') {
            imgStyle += 'width: ' + width + '; ';
        }

        if (height && height !== 'auto') {
            imgStyle += 'height: ' + height + '; ';
        }

        if (maxWidth) {
            imgStyle += 'max-width: ' + maxWidth + '; ';
        }

        if (maxHeight) {
            imgStyle += 'max-height: ' + maxHeight + '; ';
        }

        // Maintain aspect ratio if specified
        if (keepRatio) {
            imgStyle += 'object-fit: contain; ';
        }

        // Container style for alignment
        var containerStyle = '';
        var imgDisplay = 'block';

        if (align === 'center') {
            containerStyle = 'text-align: center; ';
            imgDisplay = 'inline-block';
        } else if (align === 'right') {
            containerStyle = 'text-align: right; ';
        } else if (align === 'left') {
            containerStyle = 'text-align: left; ';
        }

        // For float positioning
        if (align === 'float-left') {
            imgStyle += 'float: left; margin: 0 15px 10px 0; ';
        } else if (align === 'float-right') {
            imgStyle += 'float: right; margin: 0 0 10px 15px; ';
        }

        // For inline positioning (stays on same line, no overflow)
        var isInline = false;
        if (align === 'inline-right') {
            containerStyle = 'display: inline-flex; align-items: center; flex-shrink: 0; margin-left: auto; margin-right: 36px; ';
            imgStyle += 'vertical-align: middle; ';
            imgDisplay = 'block';
            isInline = true;
        } else if (align === 'inline-left') {
            containerStyle = 'display: inline-flex; align-items: center; flex-shrink: 0; ';
            imgStyle += 'vertical-align: middle; ';
            imgDisplay = 'block';
            isInline = true;
        }

        // Only add block margin if not inline
        var wrapperMargin = isInline ? '' : 'margin: 10px 0;';
        var html = '<div style="' + containerStyle + wrapperMargin + '">';
        html += '<img src="' + src + '" alt="' + alt + '" style="' + imgStyle + 'display: ' + imgDisplay + ';" />';
        html += '</div>';

        return html;
    }

    function processTableTag(content, data, numDimensions, numMeasures, layout, headerColor, stripedRows, hideHeader, font, size, style, align, sort, widths, itemMapping) {
        var lines = content.trim().split('\n');
        itemMapping = itemMapping || {};
        if (lines.length === 0 || !data || !data.qMatrix) {
            return '<span class="no-data">No data</span>';
        }

        // Default colors and alignment
        headerColor = headerColor || '#009845';
        stripedRows = stripedRows !== 'false';

        // Parse per-column alignment (comma-separated: "left,right,center")
        var alignments = [];
        if (align && align.indexOf(',') >= 0) {
            alignments = align.split(',').map(function(a) { return a.trim(); });
        } else {
            // Single alignment for all columns
            alignments = [align || 'left'];
        }

        // Parse manual column widths if provided (comma-separated percentages)
        var manualWidths = null;
        if (widths) {
            var rawWidths = widths.split(',').map(function(w) {
                return parseFloat(w.trim());
            });

            // Normalize to 100%
            var total = rawWidths.reduce(function(a, b) { return a + b; }, 0);
            manualWidths = rawWidths.map(function(w) {
                return Math.round((w / total) * 100) + '%';
            });

        }

        var tableStyle = 'width: 100%; border-collapse: collapse; margin: 0; table-layout: fixed; box-sizing: border-box;';
        var tableHtml = '<table class="jz-table-debug" style="' + tableStyle + '">';

        // Parse columns
        var columns = [];
        lines.forEach(function(line, idx) {
            var parts = line.split('|');
            if (parts.length >= 2) {
                var header = parts[0].trim();
                var placeholder = parts[1].trim();
                var colAlign = alignments[idx] || alignments[0] || 'left';
                columns.push({ header: header, placeholder: placeholder, align: colAlign });
            }
        });

        // Calculate column widths
        var columnWidths = [];


        if (manualWidths && manualWidths.length === columns.length) {
            // Use manual widths if provided and count matches
            columnWidths = manualWidths;
        } else {
            if (manualWidths) {
                // console.warn('[TABLE COLUMN WIDTHS] Manual widths count (' + manualWidths.length + ') does not match columns count (' + columns.length + '). Using smart calculation instead.');
            }

            // Smart calculation based on actual content analysis
            var maxContentLengths = [];
            var columnTypes = []; // 'numeric', 'text', or 'mixed'

            // Analyze header lengths
            columns.forEach(function(col, idx) {
                maxContentLengths[idx] = col.header ? col.header.length : 5; // Minimum 5 for empty
                columnTypes[idx] = { numericCount: 0, textCount: 0 };
            });

            // Analyze data content lengths AND types (sample up to 50 rows for performance)
            var sampleSize = Math.min(50, data.qMatrix.length);
            for (var i = 0; i < sampleSize; i++) {
                var row = data.qMatrix[i];
                columns.forEach(function(col, idx) {
                    var value = getPlaceholderValue(col.placeholder, row, numDimensions, numMeasures, itemMapping);
                    if (value) {
                        var strValue = String(value);
                        var len = strValue.length;
                        maxContentLengths[idx] = Math.max(maxContentLengths[idx], len);

                        // Detect if content is numeric (numbers, commas, periods, currency symbols)
                        var cleanValue = strValue.replace(/[$€£¥,\s]/g, '');
                        var isNumeric = !isNaN(cleanValue) && cleanValue.length > 0;

                        if (isNumeric) {
                            columnTypes[idx].numericCount++;
                        } else {
                            columnTypes[idx].textCount++;
                        }
                    }
                });
            }

            // Determine column type (numeric if 80%+ are numeric)
            var finalColumnTypes = columnTypes.map(function(type) {
                var total = type.numericCount + type.textCount;
                if (total === 0) return 'text'; // Default to text if no data
                var numericPercent = type.numericCount / total;
                return numericPercent >= 0.8 ? 'numeric' : 'text';
            });

            // Calculate weights based on content length AND type
            var totalWeight = 0;
            var weights = maxContentLengths.map(function(len, idx) {
                var type = finalColumnTypes[idx];

                // Base weight from content length (use square root to prevent dominance)
                var baseWeight = Math.max(1, Math.sqrt(len * 2));

                // Apply type multiplier:
                // - Numeric columns: 1.0x (normal weight - formatted numbers need space)
                // - Text columns: 1.5x (make them wider)
                var typeMultiplier = type === 'numeric' ? 1.0 : 1.5;
                var weight = baseWeight * typeMultiplier;

                totalWeight += weight;
                return weight;
            });

            // Convert weights to percentages
            var rawPercentages = weights.map(function(w) {
                return (w / totalWeight) * 100;
            });

            // Apply minimum width for numeric columns (need space for formatted numbers)
            var adjustedPercentages = rawPercentages.map(function(p, idx) {
                var type = finalColumnTypes[idx];
                if (type === 'numeric') {
                    return Math.max(p, 15); // Minimum 15% for numeric columns
                }
                return p;
            });

            // Renormalize to 100%
            var percentTotal = adjustedPercentages.reduce(function(a, b) { return a + b; }, 0);
            columnWidths = adjustedPercentages.map(function(p) {
                return Math.round((p / percentTotal) * 100) + '%';
            });

        }

        // Store width percentages as data attribute for later pixel calculation
        var widthPercentages = columnWidths.map(function(w) { return parseFloat(w); }).join(',');

        // Replace opening table tag with data attribute
        tableHtml = tableHtml.replace('<table class="jz-table-debug"', '<table class="jz-table-debug" data-col-percentages="' + widthPercentages + '"');

        // Add colgroup to enforce widths with table-layout: fixed
        tableHtml += '<colgroup>';
        columnWidths.forEach(function(width) {
            tableHtml += '<col style="width: ' + width + '; min-width: 0; box-sizing: border-box;">';
        });
        tableHtml += '</colgroup>';

        // Only render header if not hidden
        if (!hideHeader) {
            tableHtml += '<thead><tr>';
            columns.forEach(function(col, idx) {
                var headerStyle = 'background-color: ' + headerColor + '; color: white; padding: 3px 8px; text-align: ' + col.align + '; font-weight: bold; border: 1px solid #ddd; word-wrap: break-word; overflow-wrap: break-word;';
                tableHtml += '<th style="' + headerStyle + '">' + (col.header || '') + '</th>';
            });
            tableHtml += '</tr></thead>';
        }

        tableHtml += '<tbody>';

        // Identify dimension and measure columns
        var dimColumns = [];
        var measureColumns = [];
        columns.forEach(function(col) {
            var dimMatch = col.placeholder.match(/\{\{dim(\d+)\}\}/);
            var measureMatch = col.placeholder.match(/\{\{measure(\d+)\}\}/);
            if (dimMatch) {
                dimColumns.push({
                    column: col,
                    dimIndex: parseInt(dimMatch[1]) - 1
                });
            } else if (measureMatch) {
                measureColumns.push({
                    column: col,
                    measureIndex: parseInt(measureMatch[1]) - 1
                });
            }
        });

        // If table has dimensions, aggregate rows by dimension values
        var aggregatedRows = [];

        if (dimColumns.length > 0) {
            var rowMap = {};

            data.qMatrix.forEach(function(row, rowIdx) {
                // Create key from dimension values
                var key = dimColumns.map(function(dc) {
                    var val = row[dc.dimIndex] ? row[dc.dimIndex].qText : '';
                    return val;
                }).join('|');

                if (rowIdx < 3) {
                }

                if (!rowMap[key]) {
                    rowMap[key] = {
                        dimValues: {},
                        measureSums: {},
                        measureCounts: {},
                        originalRow: row
                    };
                    // Store dimension values
                    dimColumns.forEach(function(dc) {
                        rowMap[key].dimValues[dc.dimIndex] = row[dc.dimIndex];
                    });
                    // Initialize measure sums
                    measureColumns.forEach(function(mc) {
                        rowMap[key].measureSums[mc.measureIndex] = 0;
                        rowMap[key].measureCounts[mc.measureIndex] = 0;
                    });
                }

                // Add measure values
                measureColumns.forEach(function(mc) {
                    var cellIdx = numDimensions + mc.measureIndex;
                    var cell = row[cellIdx];
                    if (cell && cell.qNum !== undefined && cell.qNum !== null && !isNaN(cell.qNum)) {
                        rowMap[key].measureSums[mc.measureIndex] += cell.qNum;
                        rowMap[key].measureCounts[mc.measureIndex]++;
                    }
                });
            });

            // Convert map to array
            for (var key in rowMap) {
                aggregatedRows.push(rowMap[key]);
            }
        } else {
            // No dimensions, just show all rows
            data.qMatrix.forEach(function(row) {
                aggregatedRows.push({ originalRow: row });
            });
        }

        // Sort aggregated rows if sort parameter is provided
        if (sort) {
            var sortParts = sort.split(',');
            var sortColIndex = parseInt(sortParts[0] || '0', 10);
            var sortDirection = (sortParts[1] || 'asc').toLowerCase();

            aggregatedRows.sort(function(a, b) {
                var colInfo = columns[sortColIndex];
                if (!colInfo) return 0;

                var dimMatch = colInfo.placeholder.match(/\{\{dim(\d+)\}\}/);
                var measureMatch = colInfo.placeholder.match(/\{\{measure(\d+)\}\}/);

                var aVal, bVal;

                if (dimMatch) {
                    var dimIndex = parseInt(dimMatch[1]) - 1;
                    aVal = a.dimValues && a.dimValues[dimIndex] ? a.dimValues[dimIndex].qText : '';
                    bVal = b.dimValues && b.dimValues[dimIndex] ? b.dimValues[dimIndex].qText : '';

                    // Try numeric comparison first
                    var aNum = parseFloat(aVal);
                    var bNum = parseFloat(bVal);
                    if (!isNaN(aNum) && !isNaN(bNum)) {
                        return sortDirection === 'asc' ? aNum - bNum : bNum - aNum;
                    }
                    // Fall back to string comparison
                    return sortDirection === 'asc'
                        ? aVal.localeCompare(bVal)
                        : bVal.localeCompare(aVal);
                } else if (measureMatch) {
                    var measureIndex = parseInt(measureMatch[1]) - 1;
                    aVal = a.measureSums && a.measureSums[measureIndex] !== undefined ? a.measureSums[measureIndex] : 0;
                    bVal = b.measureSums && b.measureSums[measureIndex] !== undefined ? b.measureSums[measureIndex] : 0;
                    return sortDirection === 'asc' ? aVal - bVal : bVal - aVal;
                }

                return 0;
            });
        }

        // Populate table rows with aggregated data
        aggregatedRows.forEach(function(aggRow, idx) {
            var rowBg = stripedRows && idx % 2 === 1 ? 'rgba(0, 0, 0, 0.02)' : 'transparent';
            var rowStyle = 'background-color: ' + rowBg + ';';
            tableHtml += '<tr style="' + rowStyle + '">';

            columns.forEach(function(col, colIdx) {
                var value;
                var dimMatch = col.placeholder.match(/\{\{dim(\d+)\}\}/);
                var measureMatch = col.placeholder.match(/\{\{measure(\d+)\}\}/);

                if (dimMatch) {
                    var dimIndex = parseInt(dimMatch[1]) - 1;
                    if (aggRow.dimValues && aggRow.dimValues[dimIndex]) {
                        value = aggRow.dimValues[dimIndex].qText;
                    } else {
                        value = getPlaceholderValue(col.placeholder, aggRow.originalRow, numDimensions, numMeasures, itemMapping);
                    }
                } else if (measureMatch) {
                    var measureIndex = parseInt(measureMatch[1]) - 1;
                    if (aggRow.measureSums && aggRow.measureSums[measureIndex] !== undefined) {
                        // Format aggregated value
                        var sum = aggRow.measureSums[measureIndex];
                        var cellIdx = numDimensions + measureIndex;
                        var firstCell = aggRow.originalRow[cellIdx];
                        if (firstCell && firstCell.qText) {
                            // Detect decimal places from first cell
                            var decimalMatch = firstCell.qText.match(/\.(\d+)/);
                            var decimals = decimalMatch ? decimalMatch[1].length : 0;

                            // Format number with detected decimal places
                            var formattedNum;
                            if (decimals > 0) {
                        formattedNum = sum.toFixed(decimals).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
                            } else {
                        formattedNum = Math.round(sum).toLocaleString();
                            }

                            // Try to match the format (currency symbols, etc.)
                            var formatMatch = firstCell.qText.match(/^([^\d-]*)(-?[\d,]+\.?\d*)(.*)$/);
                            if (formatMatch) {
                        var prefix = formatMatch[1] || '';
                        var suffix = formatMatch[3] || '';
                        value = prefix + formattedNum + suffix;
                            } else {
                        value = formattedNum;
                            }
                        } else {
                            value = sum.toLocaleString();
                        }
                    } else {
                        value = getPlaceholderValue(col.placeholder, aggRow.originalRow, numDimensions, numMeasures, itemMapping);
                    }
                } else {
                    value = getPlaceholderValue(col.placeholder, aggRow.originalRow, numDimensions, numMeasures, itemMapping);
                }

                var cellStyle = 'padding: 3px 8px; border: 1px solid #ddd; text-align: ' + col.align + '; word-wrap: break-word; overflow-wrap: break-word;';
                tableHtml += '<td style="' + cellStyle + '">' + value + '</td>';
            });
            tableHtml += '</tr>';
        });

        tableHtml += '</tbody></table>';
        return wrapWithFontStyle(tableHtml, font, size, style);
    }

    function processListTag(content, data, numDimensions, numMeasures, layout, type, columns, numbering, dividers, font, size, style, align, sort, colorByDimension, itemMapping) {
        var placeholder = content.trim();
        var values = [];
        align = align || 'left';

        colorByDimension = colorByDimension || null;
        itemMapping = itemMapping || {};

        
        
        
        
        

        if (!data || !data.qMatrix) {
            return '<span class="no-data">No data</span>';
        }

        // Check if placeholder contains multiple {{...}} patterns
        var placeholderMatches = placeholder.match(/\{\{[^}]+\}\}/g);
        var hasMultiplePlaceholders = placeholderMatches && placeholderMatches.length > 1;

        // Collect all unique values with color info
        var seenValues = {};
        data.qMatrix.forEach(function(row, idx) {
            var value, dedupeValue;

            if (hasMultiplePlaceholders) {
                // Use first placeholder for deduplication
                dedupeValue = getPlaceholderValue(placeholderMatches[0], row, numDimensions, numMeasures, itemMapping);

                // Process all placeholders and combine
                var processedParts = [];
                for (var i = 0; i < placeholderMatches.length; i++) {
                    var partValue = getPlaceholderValue(placeholderMatches[i], row, numDimensions, numMeasures, itemMapping);
                    if (partValue && partValue !== 'N/A') {
                        processedParts.push(partValue);
                    }
                }

                // Replace placeholders in template with actual values
                value = placeholder;
                for (var j = 0; j < placeholderMatches.length; j++) {
                    var partVal = getPlaceholderValue(placeholderMatches[j], row, numDimensions, numMeasures, itemMapping);
                    value = value.replace(placeholderMatches[j], partVal || '');
                }
            } else {
                // Single placeholder - use existing logic
                value = getPlaceholderValue(placeholder, row, numDimensions, numMeasures, itemMapping);
                dedupeValue = value;
            }

            if (dedupeValue && dedupeValue !== 'N/A' && !seenValues[dedupeValue]) {
                seenValues[dedupeValue] = true; // Mark as seen
                try {
                    var colorInfo = { color: null, bgColor: null };
                    var colorValue = null;

                    // If colorByDimension is specified in tag attribute, use it for coloring
                    if (colorByDimension && itemMapping[colorByDimension]) {
                        var colorDimMapping = itemMapping[colorByDimension];
                        var colorDimIndex = colorDimMapping.index;
                        colorValue = row[colorDimIndex] ? row[colorDimIndex].qText : null;

                        if (colorValue) {
                            // Check for custom color mapping first, then fall back to default palette
                            colorInfo.bgColor = getColorForDimensionValue(colorValue, 'bg', layout);
                            colorInfo.color = getColorForDimensionValue(colorValue, 'text', layout);
                        } else {

                        }
                    } else {
                        // Use existing color logic from layout settings (use first placeholder if multiple)
                        var colorCheckPlaceholder = hasMultiplePlaceholders ? placeholderMatches[0] : placeholder;
                        colorInfo = getColorForValue(colorCheckPlaceholder, row, numDimensions, numMeasures, layout);
                    }

                    values.push({
                        text: value,
                        color: colorInfo.color,
                        bgColor: colorInfo.bgColor,
                        sortValue: colorValue  // Store the color dimension value for sorting
                    });
                } catch (e) {

                    // Push value without color info if color fails
                    values.push({ text: value, color: null, bgColor: null, sortValue: null });
                }
            }
        });

        

        if (values.length === 0) {
            return '<span class="no-data">No data</span>';
        }

        // Sort values if sort parameter is provided
        if (sort) {
            var sortDirection = sort.toLowerCase();
            values.sort(function(a, b) {
                // If colorByDimension is specified, sort by the color dimension value (sortValue)
                // Otherwise sort by the displayed text
                var aValue = (colorByDimension && a.sortValue) ? a.sortValue : a.text;
                var bValue = (colorByDimension && b.sortValue) ? b.sortValue : b.text;

                // Try numeric comparison first
                var aNum = parseFloat(aValue);
                var bNum = parseFloat(bValue);
                if (!isNaN(aNum) && !isNaN(bNum)) {
                    return sortDirection === 'asc' ? aNum - bNum : bNum - aNum;
                }
                // Fall back to string comparison
                return sortDirection === 'asc'
                    ? aValue.localeCompare(bValue)
                    : bValue.localeCompare(aValue);
            });
        }

        // Multi-column layout
        if (columns > 1) {
            var gapStyle = dividers ? '1px solid #e0e0e0' : '0';

            // Split items into columns manually to control distribution
            // Left column gets more items: ceil(total/columns)
            var itemsPerColumn = Math.ceil(values.length / columns);

            var containerStyle = 'display: flex; gap: 10px;';
            var listHtml = '<div style="' + containerStyle + '">';

            for (var col = 0; col < columns; col++) {
                var startIdx = col * itemsPerColumn;
                var endIdx = Math.min(startIdx + itemsPerColumn, values.length);
                var columnItems = values.slice(startIdx, endIdx);

                var columnStyle = 'flex: 1; min-width: 0;';
                if (dividers && col > 0) {
                    columnStyle += ' border-left: ' + gapStyle + '; padding-left: 10px;';
                }

                listHtml += '<div style="' + columnStyle + '">';
                columnItems.forEach(function(item, idx) {
                    var globalIdx = startIdx + idx;
                    var itemPrefix = '';
                    if (type === 'numbered') {
                        itemPrefix = (globalIdx + 1) + '.';
                    } else if (type === 'bulleted') {
                        itemPrefix = '•';
                    }

                    var containerStyle = 'display: flex; gap: 8px; margin-bottom: 0px; line-height: 1.3; text-align: ' + align + ';';
                    var bulletStyle = 'flex-shrink: 0; min-width: 20px;';
                    var textStyle = 'flex-grow: 1;';

                    if (item.color) textStyle += ' color: ' + item.color + ';';
                    // Skip white backgrounds to keep transparent
                    if (item.bgColor && item.bgColor.toLowerCase() !== '#ffffff' && item.bgColor.toLowerCase() !== '#fff') {
                        textStyle += ' background-color: ' + item.bgColor + '; border-radius: 3px; padding: 2px 10px;';
                    }

                    listHtml += '<div style="' + containerStyle + '">';
                    listHtml += '<span style="' + bulletStyle + '">' + itemPrefix + '</span>';
                    listHtml += '<span style="' + textStyle + '">' + item.text + '</span>';
                    listHtml += '</div>';
                });
                listHtml += '</div>';
            }

            listHtml += '</div>';
            return wrapWithFontStyle(listHtml, font, size, style);
        }

        // Single column list
        var listHtml = '<div style="margin: 0; padding: 0;">';
        values.forEach(function(item, idx) {
            var itemPrefix = '';
            if (type === 'numbered') {
                itemPrefix = (idx + 1) + '.';
            } else if (type === 'bulleted') {
                itemPrefix = '•';
            } else if (type === 'plain') {
                itemPrefix = '•';
            }

            var containerStyle = 'display: flex; gap: 8px; margin-bottom: 0px; line-height: 1.3; text-align: ' + align + ';';
            var bulletStyle = 'flex-shrink: 0; min-width: 20px;';
            var textStyle = 'flex-grow: 1;';

            if (item.color) textStyle += ' color: ' + item.color + ';';
            // Skip white backgrounds to keep transparent
            if (item.bgColor && item.bgColor.toLowerCase() !== '#ffffff' && item.bgColor.toLowerCase() !== '#fff') {
                textStyle += ' background-color: ' + item.bgColor + '; border-radius: 3px; padding: 2px 10px;';
            }

            listHtml += '<div style="' + containerStyle + '">';
            listHtml += '<span style="' + bulletStyle + '">' + itemPrefix + '</span>';
            listHtml += '<span style="' + textStyle + '">' + item.text + '</span>';
            listHtml += '</div>';
        });
        listHtml += '</div>';
        return wrapWithFontStyle(listHtml, font, size, style);
    }

    function processConcatTag(content, data, numDimensions, numMeasures, layout, delimiter, font, size, style, itemMapping) {
        var placeholder = content.trim();
        var values = [];
        itemMapping = itemMapping || {};

        if (!data || !data.qMatrix) {
            return '<span class="no-data">No data</span>';
        }

        // Collect all unique values
        var seenValues = {};
        data.qMatrix.forEach(function(row) {
            var value = getPlaceholderValue(placeholder, row, numDimensions, numMeasures, itemMapping);
            if (value && value !== 'N/A' && !seenValues[value]) {
                seenValues[value] = true; // Mark as seen
                values.push(value);
            }
        });

        if (values.length === 0) {
            return '<span class="no-data">No data</span>';
        }

        var result = values.join(delimiter);
        // Wrap with compact line-height for cleaner display
        var wrappedResult = '<span style="line-height: 1.3;">' + result + '</span>';
        return wrapWithFontStyle(wrappedResult, font, size, style);
    }

    function processGridTag(content, data, numDimensions, numMeasures, layout, columns, dividers, font, size, style, itemMapping) {
        var placeholder = content.trim();
        var values = [];
        itemMapping = itemMapping || {};

        if (!data || !data.qMatrix) {
            return '<span class="no-data">No data</span>';
        }

        // Collect all unique values
        var seenValues = {};
        data.qMatrix.forEach(function(row) {
            var value = getPlaceholderValue(placeholder, row, numDimensions, numMeasures, itemMapping);
            if (value && value !== 'N/A' && !seenValues[value]) {
                seenValues[value] = true; // Mark as seen
                values.push(value);
            }
        });

        if (values.length === 0) {
            return '<span class="no-data">No data</span>';
        }

        var gapStyle = dividers ? '0' : '15px';
        var gridHtml = '<div style="display: grid; grid-template-columns: repeat(' + columns + ', 1fr); gap: ' + gapStyle + ';">';
        values.forEach(function(value, idx) {
            var itemStyle = 'padding: 8px; border: 1px solid #ddd; border-radius: 4px; background: transparent;';
            if (dividers) {
                itemStyle = 'padding: 8px; border: 1px solid #ddd; background: transparent;';
                // Add right border for dividers (except last column)
                var colIndex = idx % columns;
                if (colIndex < columns - 1) {
                    itemStyle += ' border-right: 2px solid #999;';
                }
            }
            gridHtml += '<div style="' + itemStyle + '">';
            gridHtml += value;
            gridHtml += '</div>';
        });
        gridHtml += '</div>';
        return wrapWithFontStyle(gridHtml, font, size, style);
    }

    function processPivotTag(content, data, numDimensions, numMeasures, layout, headerColor, stripedRows, font, size, style, align, itemMapping) {
        var lines = content.trim().split('\n');
        itemMapping = itemMapping || {};

        if (lines.length === 0) {
            return '<span class="no-data">No data</span>';
        }

        // Parse alignment (default: left for row labels, right for data columns)
        var alignments = [];
        if (align) {
            alignments = align.split(',').map(function(a) { return a.trim(); });
        }

        // Parse structure to determine column count
        var firstLine = lines[0].trim();
        var hasHeaders = firstLine.startsWith('|');
        var numDataColumns = 0;
        var headers = [];

        if (hasHeaders) {
            headers = firstLine.substring(1).split('|').map(function(h) { return h.trim(); });
            numDataColumns = headers.length;
        } else if (lines.length > 0) {
            // Count from first data row
            var parts = lines[0].split('|');
            numDataColumns = parts.length - 1; // Subtract 1 for row label
        }

        // Calculate smart column widths
        var totalColumns = 1 + numDataColumns; // 1 for row labels + N for data
        var columnWidths = [];

        // Analyze content to determine optimal widths
        var maxContentLengths = [];

        // Initialize with row label column
        maxContentLengths[0] = 10; // Default for row label column

        // Initialize data columns with header lengths if present
        if (hasHeaders) {
            headers.forEach(function(header, idx) {
                maxContentLengths[idx + 1] = header.length;
            });
        } else {
            for (var c = 1; c < totalColumns; c++) {
                maxContentLengths[c] = 8;
            }
        }

        // Analyze actual data content lengths (sample all rows in pivot since they're typically few)
        var startIndex = hasHeaders ? 1 : 0;
        for (var i = startIndex; i < lines.length; i++) {
            var line = lines[i].trim();
            if (!line) continue;

            var parts = line.split('|');
            var rowLabel = parts[0].trim();
            var cells = parts.slice(1);

            // Row label length
            maxContentLengths[0] = Math.max(maxContentLengths[0], rowLabel.length);

            // Data cell lengths
            cells.forEach(function(cell, idx) {
                var cellContent = cell.trim();

                // Get the actual value length after placeholder replacement
                if (data && data.qMatrix && data.qMatrix.length > 0) {
                    var firstRow = data.qMatrix[0];
                    var replacedContent = cellContent.replace(/\{\{[^\}]+\}\}/g, function(placeholder) {
                        return getPlaceholderValue(placeholder, firstRow, numDimensions, numMeasures, itemMapping) || '';
                    });
                    maxContentLengths[idx + 1] = Math.max(maxContentLengths[idx + 1] || 0, replacedContent.length);
                } else {
                    maxContentLengths[idx + 1] = Math.max(maxContentLengths[idx + 1] || 0, cellContent.length);
                }
            });
        }

        // Calculate weights (use square root to prevent dominance)
        var totalWeight = 0;
        var weights = maxContentLengths.map(function(len) {
            var weight = Math.max(1, Math.sqrt(len * 2));
            totalWeight += weight;
            return weight;
        });

        // Convert to percentages
        columnWidths = weights.map(function(w) {
            return Math.round((w / totalWeight) * 100) + '%';
        });

        // Build table with colgroup for column widths
        var tableStyle = 'width: 100%; border-collapse: collapse; margin: 0; table-layout: fixed; box-sizing: border-box;';
        var tableHtml = '<table class="jz-table-debug" style="' + tableStyle + '">';

        // Store percentage widths as data attribute for post-render pixel conversion
        var widthPercentages = columnWidths.map(function(w) { return parseFloat(w); }).join(',');
        tableHtml = tableHtml.replace('<table class="jz-table-debug"', '<table class="jz-table-debug" data-col-percentages="' + widthPercentages + '"');

        // Add colgroup to define column widths
        tableHtml += '<colgroup>';
        columnWidths.forEach(function(width) {
            tableHtml += '<col style="width: ' + width + '; min-width: 0; box-sizing: border-box;">';
        });
        tableHtml += '</colgroup>';

        // Process first row as column headers (starts with |)
        if (hasHeaders) {
            tableHtml += '<thead><tr>';
            // Empty cell for row label column
            tableHtml += '<th style="background-color: ' + headerColor + '; color: white; padding: 3px 8px; text-align: left; font-weight: bold; border: 1px solid #ddd;"></th>';
            // Column headers
            headers.forEach(function(header, idx) {
                var colAlign = alignments[idx + 1] || alignments[1] || 'right'; // Default right for data columns
                tableHtml += '<th style="background-color: ' + headerColor + '; color: white; padding: 3px 8px; text-align: ' + colAlign + '; font-weight: bold; border: 1px solid #ddd;">' + header + '</th>';
            });
            tableHtml += '</tr></thead>';
        }

        // Process data rows
        tableHtml += '<tbody>';
        var startIndex = hasHeaders ? 1 : 0;

        for (var i = startIndex; i < lines.length; i++) {
            var line = lines[i].trim();
            if (!line) continue;

            var parts = line.split('|');
            var rowLabel = parts[0].trim();
            var cells = parts.slice(1);

            var rowBg = stripedRows && (i - startIndex) % 2 === 1 ? 'rgba(0, 0, 0, 0.02)' : 'transparent';
            tableHtml += '<tr style="background-color: ' + rowBg + ';">';

            // Row label (first column)
            var rowAlign = alignments[0] || 'left';
            tableHtml += '<td style="padding: 3px 8px; border: 1px solid #ddd; text-align: ' + rowAlign + '; font-weight: 600;">' + rowLabel + '</td>';

            // Data cells - replace master item placeholders with actual values
            cells.forEach(function(cell, idx) {
                var cellContent = cell.trim();

                // Replace master item placeholders with actual values from the data
                if (!data || !data.qMatrix || data.qMatrix.length === 0) {
                    cellContent = '<span class="no-data">No data</span>';
                } else {
                    // Use first row for pivot table values
                    var firstRow = data.qMatrix[0];
                    cellContent = cellContent.replace(/\{\{[^\}]+\}\}/g, function(placeholder) {
                        return getPlaceholderValue(placeholder, firstRow, numDimensions, numMeasures, itemMapping);
                    });
                }

                var colAlign = alignments[idx + 1] || alignments[1] || 'right';
                tableHtml += '<td style="padding: 3px 8px; border: 1px solid #ddd; text-align: ' + colAlign + ';">' + cellContent + '</td>';
            });

            tableHtml += '</tr>';
        }

        tableHtml += '</tbody></table>';
        return wrapWithFontStyle(tableHtml, font, size, style);
    }

    function processKpiTag(content, data, numDimensions, numMeasures, layout, label, font, size, style, itemMapping) {
        var placeholder = content.trim();
        itemMapping = itemMapping || {};

        if (!data || !data.qMatrix || data.qMatrix.length === 0) {
            return '<span class="no-data">No data</span>';
        }

        // Get first value only for KPI
        var value = getPlaceholderValue(placeholder, data.qMatrix[0], numDimensions, numMeasures, itemMapping);

        var kpiHtml = '<div style="text-align: center; padding: 8px;">';
        if (label) {
            kpiHtml += '<div style="font-size: 13px; color: #666; margin-bottom: 5px;">' + label + '</div>';
        }
        kpiHtml += '<div style="font-size: 32px; font-weight: bold; color: #009845;">' + value + '</div>';
        kpiHtml += '</div>';
        return wrapWithFontStyle(kpiHtml, font, size, style);
    }

    function processBoxTag(content, color, bgColor, data, numDimensions, numMeasures, layout, font, size, style, align) {
        // Default styling for box
        var defaultBgColor = bgColor || 'transparent';
        var borderColor = color || '#009845';
        var textAlign = align || 'left';

        var boxStyle = 'border-left: 4px solid ' + borderColor + '; ' +
                      'background-color: ' + defaultBgColor + '; ' +
                      'padding: 8px; ' +
                      'margin: 3px 0; ' +
                      'border-radius: 4px; ' +
                      'box-shadow: 0 2px 4px rgba(0,0,0,0.1); ' +
                      'text-align: ' + textAlign + ';';

        // Recursively process any nested content tags first
        var processedContent = processContentTags(content, data, numDimensions, numMeasures, layout, null, {});

        // Always parse markdown (safe to run on mixed HTML/markdown content)
        var html = parseMarkdown(processedContent);

        // Apply font styling to content if specified
        boxStyle = applyFontStyle(boxStyle, font, size, style);

        return '<div style="' + boxStyle + '">' + html + '</div>';
    }

    function processVizTag(vizId, vizName, height) {
        if (!vizId && !vizName) {
            return '<div style="padding: 8px; background: #fff0f0; border: 1px solid #d32f2f; border-radius: 4px; color: #d32f2f;">Error: No visualization ID or name specified. Use: #[viz id="your-viz-id"]#[/viz] or #[viz name="Trend"]#[/viz]</div>';
        }

        // Generate unique container ID
        var identifier = vizId || vizName;
        var containerId = 'qv-viz-' + identifier.replace(/[^a-zA-Z0-9]/g, '-') + '-' + Math.random().toString(36).substr(2, 9);

        var containerStyle = 'width: 100%; height: ' + height + '; margin: 10px 0; border: 1px solid #ddd; border-radius: 4px; overflow: hidden;';

        // Store viz info in data attributes for later rendering
        var html = '<div id="' + containerId + '" class="qlik-viz-container"';
        if (vizId) {
            html += ' data-viz-id="' + vizId + '"';
        }
        if (vizName) {
            html += ' data-viz-name="' + vizName + '"';
        }
        html += ' style="' + containerStyle + '"></div>';

        return html;
    }

    function processClaudeTag(prompt, dataRefs, content, data, numDimensions, numMeasures, layout) {
        // Check if Claude AI is enabled
        if (CLAUDE_ENABLED !== 1) {
            return '<div style="padding: 8px; background: #fff3cd; border-left: 4px solid #ffc107; border-radius: 4px; margin: 5px 0; color: #856404;">⚠️ Claude AI features are disabled.</div>';
        }

        // Check if user has permission to use Claude
        if (!userHasFeaturePermission('jz_claude')) {
            return '<div style="padding: 8px; background: #fff3cd; border-left: 4px solid #ffc107; border-radius: 4px; margin: 5px 0; color: #856404;">⚠️ You do not have permission to use Claude AI features.</div>';
        }

        // Generate unique container ID
        var containerId = 'claude-analysis-' + Math.random().toString(36).substr(2, 9);

        // Extract data based on references
        var extractedData = extractDataForClaude(dataRefs, data, numDimensions, numMeasures);

        var containerStyle = 'padding: 8px; background: #f8f9fa; border-left: 4px solid #1976D2; border-radius: 4px; margin: 5px 0;';

        var html = '<div id="' + containerId + '" class="claude-analysis-container" style="' + containerStyle + '"';
        html += ' data-prompt="' + prompt.replace(/"/g, '&quot;') + '"';
        html += ' data-extracted-data="' + encodeURIComponent(JSON.stringify(extractedData)) + '"';
        html += ' data-orchestrator-url="' + (layout.mcpOrchestratorUrl || 'https://gse-mcp.replit.app') + '"';
        html += ' data-endpoint="/api/execute-tool"';
        html += ' data-system-prompt="' + (layout.claudeSystemPrompt || 'You are a business intelligence analyst. Answer directly with only the requested information. Do not include preambles like \'Here are\' or \'Based on the data\'. Just provide the answer.').replace(/"/g, '&quot;') + '"';
        html += ' data-max-tokens="' + (layout.claudeMaxTokens || 2000) + '"';
        html += '>';
        html += '<div class="claude-loading" style="text-align: center; color: #666;">';
        html += '<div style="display: inline-block; width: 20px; height: 20px; border: 3px solid #ddd; border-top-color: #1976D2; border-radius: 50%; animation: spin 1s linear infinite;"></div>';
        html += '<p style="margin-top: 10px;">🤖 Analyzing data with Claude AI...</p>';
        html += '</div>';
        html += '</div>';

        return html;
    }

    function extractDataForClaude(dataRefs, data, numDimensions, numMeasures) {
        if (!data || !data.qMatrix) {
            return { error: 'No data available' };
        }

        var result = {
            dimensions: {},
            measures: {},
            rows: []
        };

        // Parse data references (e.g., "{{dim1}},{{measure1}}")
        var refs = dataRefs ? dataRefs.split(',').map(function(r) { return r.trim(); }) : [];

        // If no specific refs, include all data
        if (refs.length === 0) {
            // Include all dimensions and measures
            data.qMatrix.forEach(function(row, idx) {
                var rowData = {};
                for (var d = 0; d < numDimensions; d++) {
                    rowData['dim' + (d + 1)] = row[d] ? row[d].qText : null;
                }
                for (var m = 0; m < numMeasures; m++) {
                    rowData['measure' + (m + 1)] = row[numDimensions + m] ? row[numDimensions + m].qText : null;
                }
                result.rows.push(rowData);
            });
        } else {
            // Extract only specified data
            data.qMatrix.forEach(function(row) {
                var rowData = {};
                refs.forEach(function(ref) {
                    var dimMatch = ref.match(/\{\{dim(\d+)\}\}/);
                    var measureMatch = ref.match(/\{\{measure(\d+)\}\}/);

                    if (dimMatch) {
                        var dimIdx = parseInt(dimMatch[1]) - 1;
                        if (dimIdx < numDimensions && row[dimIdx]) {
                            rowData['dim' + (dimIdx + 1)] = row[dimIdx].qText;
                        }
                    } else if (measureMatch) {
                        var measureIdx = parseInt(measureMatch[1]) - 1;
                        var cellIdx = numDimensions + measureIdx;
                        if (measureIdx < numMeasures && row[cellIdx]) {
                            rowData['measure' + (measureIdx + 1)] = row[cellIdx].qText;
                        }
                    }
                });
                result.rows.push(rowData);
            });
        }

        return result;
    }

    function getPlaceholderValue(placeholder, row, numDimensions, numMeasures, itemMapping) {
        // Check for dimension placeholder {{dim1}}, {{dim2}}, etc.
        var dimMatch = placeholder.match(/\{\{dim(\d+)\}\}/);
        if (dimMatch) {
            var dimIdx = parseInt(dimMatch[1]) - 1;
            if (dimIdx < numDimensions && row[dimIdx]) {
                return row[dimIdx].qText;
            }
        }

        // Check for measure placeholder {{measure1}}, {{measure2}}, etc.
        var measureMatch = placeholder.match(/\{\{measure(\d+)\}\}/);
        if (measureMatch) {
            var measureIdx = parseInt(measureMatch[1]) - 1;
            var cellIdx = numDimensions + measureIdx;
            if (measureIdx < numMeasures && row[cellIdx]) {
                return row[cellIdx].qText;
            }
        }

        // Check for master item placeholder {{[Master Item]}}
        var masterMatch = placeholder.match(/\{\{\[([^\]]+)\]\}\}/);
        if (masterMatch && itemMapping) {
            var itemName = masterMatch[1];
            var mapping = itemMapping[itemName];
            if (mapping) {
                var idx = mapping.type === 'dim' ? mapping.index : numDimensions + mapping.index;
                if (row[idx]) {
                    return row[idx].qText;
                }
            }
        }

        return 'N/A';
    }

    // Get color and background color for a dimension value based on layout settings
    function getColorForValue(placeholder, row, numDimensions, numMeasures, layout) {
        var result = { color: null, bgColor: null };

        // Safety checks
        if (!row || !layout) return result;

        // Check which dimension this placeholder refers to
        var dimMatch = placeholder.match(/\{\{dim(\d+)\}\}/);
        if (!dimMatch) return result;

        var targetDimIdx = parseInt(dimMatch[1]) - 1;
        if (targetDimIdx >= numDimensions || targetDimIdx < 0) return result;
        if (!row[targetDimIdx]) return result;

        var targetValue = row[targetDimIdx].qText;
        if (!targetValue) return result;

        // NEW: Check for "By dimension" color mode
        var colorModeField = 'dim' + (targetDimIdx + 1) + 'ColorMode';
        var colorDimensionField = 'dim' + (targetDimIdx + 1) + 'ColorDimension';
        var useLibraryColorsField = 'dim' + (targetDimIdx + 1) + 'UseLibraryColors';

        if (layout[colorModeField] === 'byDimension' && layout[colorDimensionField]) {
            // Apply palette color based on the dimension value
            result.color = getDefaultPaletteColor(targetValue);
            result.bgColor = getDefaultPaletteBgColor(targetValue);
            return result;
        }

        // OLD: Check if this dimension has color-by configuration
        var colorByField = 'dim' + (targetDimIdx + 1) + 'ColorBy';
        var bgColorByField = 'dim' + (targetDimIdx + 1) + 'BgColorBy';
        var colorBy = layout[colorByField];
        var bgColorBy = layout[bgColorByField];

        // Get foreground color
        if (colorBy) {
            var colorDimMatch = colorBy.match(/dim(\d+)/);
            if (colorDimMatch) {
                var colorDimIdx = parseInt(colorDimMatch[1]) - 1;
                if (colorDimIdx < numDimensions && row[colorDimIdx]) {
                    var colorDimValue = row[colorDimIdx].qText;
                    // Check for custom color mapping or use attribute color
                    if (row[colorDimIdx].qAttrExps && row[colorDimIdx].qAttrExps.qValues &&
                        row[colorDimIdx].qAttrExps.qValues[0] && row[colorDimIdx].qAttrExps.qValues[0].qText) {
                        result.color = row[colorDimIdx].qAttrExps.qValues[0].qText;
                    } else {
                        // Use default palette
                        result.color = getDefaultPaletteColor(colorDimValue);
                    }
                }
            }
        }

        // Get background color
        if (bgColorBy) {
            var bgColorDimMatch = bgColorBy.match(/dim(\d+)/);
            if (bgColorDimMatch) {
                var bgColorDimIdx = parseInt(bgColorDimMatch[1]) - 1;
                if (bgColorDimIdx < numDimensions && row[bgColorDimIdx]) {
                    var bgColorDimValue = row[bgColorDimIdx].qText;
                    // Check for custom color mapping or use attribute color
                    if (row[bgColorDimIdx].qAttrExps && row[bgColorDimIdx].qAttrExps.qValues &&
                        row[bgColorDimIdx].qAttrExps.qValues[0] && row[bgColorDimIdx].qAttrExps.qValues[0].qText) {
                        result.bgColor = row[bgColorDimIdx].qAttrExps.qValues[0].qText;
                    } else {
                        // Use default palette (lighter for backgrounds)
                        result.bgColor = getDefaultPaletteBgColor(bgColorDimValue);
                    }
                }
            }
        }

        return result;
    }

    /**
     * Get color for a dimension value, checking custom mappings first
     * @param {string} value - The dimension value to get color for
     * @param {string} colorType - Either 'text' or 'bg'
     * @param {object} layout - The extension layout object
     * @returns {string} - The color hex code
     */
    function getColorForDimensionValue(value, colorType, layout) {
        if (!value || typeof value !== 'string') {
            return colorType === 'bg' ? getDefaultPaletteBgColor(value) : getDefaultPaletteColor(value);
        }

        // Check for custom color mappings first
        if (layout && layout.colorMappings && Array.isArray(layout.colorMappings)) {
            var mapping = layout.colorMappings.find(function(m) {
                return m.dimensionValue === value;
            });

            if (mapping) {
                if (colorType === 'bg' && mapping.bgColor) {
                    var bgColor = mapping.bgColor;
                    // Handle color-picker object format {color: "#hex"}
                    if (typeof bgColor === 'object' && bgColor.color) {
                        bgColor = bgColor.color;
                    }
                    return bgColor;
                } else if (colorType === 'text' && mapping.textColor) {
                    var textColor = mapping.textColor;
                    // Handle color-picker object format {color: "#hex"}
                    if (typeof textColor === 'object' && textColor.color) {
                        textColor = textColor.color;
                    }
                    return textColor;
                }
            }
        }

        // Fall back to default palette
        return colorType === 'bg' ? getDefaultPaletteBgColor(value) : getDefaultPaletteColor(value);
    }

    function getDefaultPaletteColor(value) {
        // Safety check for defaultPalette
        if (!defaultPalette || !Array.isArray(defaultPalette) || defaultPalette.length === 0) {
            
            return '#009845'; // Fallback color
        }

        if (!value || typeof value !== 'string') {
            
            return defaultPalette[0];
        }

        var hash = 0;
        for (var i = 0; i < value.length; i++) {
            hash = value.charCodeAt(i) + ((hash << 5) - hash);
        }
        var index = Math.abs(hash) % defaultPalette.length;
        return defaultPalette[index];
    }

    function getDefaultPaletteBgColor(value) {
        // Safety check for defaultPalette
        if (!defaultPalette || !Array.isArray(defaultPalette) || defaultPalette.length === 0) {
            
            return '#00984520'; // Fallback color with transparency
        }

        if (!value || typeof value !== 'string') {
            
            return defaultPalette[0] + '20';
        }

        var color = getDefaultPaletteColor(value);
        // Convert to lighter background version (add transparency)
        return color + '20'; // 20% opacity
    }

    function callClaudeAnalysis($container, prompt, extractedData, orchestratorUrl, endpoint, systemPrompt, maxTokens) {
        // Create cache key from prompt and data
        var cacheKey = hashData({
            prompt: prompt,
            data: extractedData.rows,
            systemPrompt: systemPrompt
        });

        // Check cache first
        if (claudeResponseCache[cacheKey]) {
            $container.html(claudeResponseCache[cacheKey]);
            return;
        }

        // Format data for Claude
        var dataText = '';
        if (extractedData.rows && extractedData.rows.length > 0) {
            dataText = 'Data:\n' + JSON.stringify(extractedData.rows, null, 2);
        } else if (extractedData.error) {
            dataText = 'Error: ' + extractedData.error;
        }

        var fullPrompt = prompt + '\n\n' + dataText;

        // Build headers
        var headers = {
            'Content-Type': 'application/json'
        };

        // Build fetch options with SSO credentials
        var fetchOptions = {
            method: 'POST',
            headers: headers,
            credentials: 'include', // SSO mode - always include cookies
            body: JSON.stringify({
                serverId: 'claude-server',
                toolName: 'claude-prompt',
                parameters: {
                    prompt: fullPrompt,
                    system_prompt: systemPrompt,
                    max_tokens: maxTokens
                }
            })
        };

        // Make API call to orchestrator using MCP server format
        fetch(orchestratorUrl + endpoint, fetchOptions)
        .then(function(response) {
            if (!response.ok) {
                throw new Error('HTTP ' + response.status + ': ' + response.statusText);
            }
            return response.json();
        })
        .then(function(data) {

            // Parse response from MCP server format
            var analysis = '';
            if (data.result && data.result.content && data.result.content[0] && data.result.content[0].text) {
                analysis = data.result.content[0].text;
            } else if (data.analysis) {
                analysis = data.analysis;
            } else if (data.response) {
                analysis = data.response;
            } else {
                analysis = 'No response received';
            }

            // Render the analysis
            var renderedHtml = '<div style="white-space: pre-wrap; line-height: 1.6;">' +
                '<div style="display: flex; align-items: center; margin-bottom: 10px; padding-bottom: 10px; border-bottom: 2px solid #1976D2;">' +
                '<strong style="font-size: 16px;">🤖 Claude AI Analysis</strong>' +
                '</div>' +
                parseMarkdown(analysis) +
                '</div>';

            // Store in cache
            claudeResponseCache[cacheKey] = renderedHtml;

            // Update container with analysis
            $container.html(renderedHtml);
        })
        .catch(function(error) {
            
            $container.html('<div style="padding: 8px; background: #fff0f0; border-left: 4px solid #d32f2f; border-radius: 4px;">' +
                '<strong style="color: #d32f2f;">❌ Error</strong><br>' +
                '<p style="margin: 10px 0; color: #666;">Failed to get Claude AI analysis:</p>' +
                '<p style="color: #999; font-family: monospace; font-size: 12px;">' + error.message + '</p>' +
                '<p style="margin-top: 10px; color: #666; font-size: 13px;">Check console for details and verify orchestrator URL in extension properties.</p>' +
                '</div>');
        });
    }

    // Section Editing Functions
    function getAppSpaceAndConnection(callback) {
        var app = qlik.currApp();

        app.getAppLayout().then(function(appLayoutModel) {
            // The model has a .layout property with the actual layout data
            var appLayout = appLayoutModel.layout || appLayoutModel;

            var appId = appLayout.qFileName;

            // Try multiple ways to get space ID
            var spaceId = null;
            if (appLayout.qMeta && appLayout.qMeta.space) {
                spaceId = appLayout.qMeta.space;
            } else if (appLayout.qMeta && appLayout.qMeta.spaceId) {
                spaceId = appLayout.qMeta.spaceId;
            } else if (appLayout.spaceId) {
                spaceId = appLayout.spaceId;
            } else if (appLayout.space) {
                spaceId = appLayout.space;
            }

            if (!spaceId) {
                
                
                callback(null);
                return;
            }

            // Get DataFiles connection for this space
            var apiUrl = 'https://jochemz.eu.qlikcloud.com/api/v1/data-connections?spaceId=' + spaceId;
            var apiKey = currentModel && currentModel.layout && currentModel.layout.apiKey ? currentModel.layout.apiKey.trim() : '';

            // 

            fetch(apiUrl, {
                headers: {
                    'Authorization': 'Bearer ' + apiKey
                }
            })
            .then(function(response) { return response.json(); })
            .then(function(data) {
                var dataFilesConnection = null;
                if (data.data) {
                    data.data.forEach(function(conn) {
                        // 
                        if (conn.qName === 'DataFiles') {
                            dataFilesConnection = conn;
                        }
                    });
                }

                if (!dataFilesConnection) {
                    
                }

                callback({
                    appId: appId,
                    spaceId: spaceId,
                    connectionId: dataFilesConnection ? dataFilesConnection.id : null
                });
            })
            .catch(function(error) {
                
                callback(null);
            });
        }).catch(function(error) {
            
            callback(null);
        });
    }

    // Load saved content for a SPECIFIC PK value (one file per PK)
    function loadModifiedSections(appId, pkValue, callback) {
        if (!pkValue) {
            callback(null);
            return;
        }

        // Check cache first - if we have fresh data for this PK, use it
        if (modifiedSectionsCache.loaded &&
            modifiedSectionsCache.pkValue === pkValue &&
            modifiedSectionsCache.data) {
            // Return cached data immediately
            callback(modifiedSectionsCache.data);
            return;
        }

        // If SharePoint is marked unavailable and hasn't recovered, skip the call
        if (sharepointUnavailable && sharepointUnavailableSince) {
            var minutesSince = (Date.now() - sharepointUnavailableSince) / 1000 / 60;
            // Try again after 10 minutes
            if (minutesSince < 10) {
                callback(null);
                return;
            } else {
                // Reset and try again

                sharepointUnavailable = false;
                sharepointUnavailableSince = null;
            }
        }

        // Sanitize pkValue for filename (replace special chars with underscores)
        var safePkValue = pkValue.replace(/[^a-zA-Z0-9-_]/g, '_');
        var fileName = 'qlik-sections-' + appId + '-' + safePkValue + '.json';

        // Try to get layout from currentModel, but it might not be set during initial render
        var layout = {};
        if (currentModel && currentModel.layout) {
            layout = currentModel.layout;
        }

        var mcpUrl = (layout.mcpOrchestratorUrl || 'https://gse-mcp.replit.app').replace(/\/+$/, '');
        var sitePath = layout.sharepointSitePath || '/sites/QlikTechnicalEnablement';
        var folderPath = layout.sharepointFolderPath || '/Shared Documents/Qlik200_Topsheet';

        var headers = { 'Content-Type': 'application/json' };

        var filePath = folderPath + '/' + fileName;

        // Track error count for this file
        var errorKey = fileName;
        if (!sharepointErrorCount[errorKey]) {
            sharepointErrorCount[errorKey] = 0;
        }

        // Build fetch options with SSO credentials
        var requestBody = {
            sitePath: sitePath,
            filePath: filePath
        };

        var fetchOptions = {
            method: 'POST',
            headers: headers,
            credentials: 'include', // SSO mode - always include cookies
            body: JSON.stringify(requestBody)
        };

        // Debug logging (only enable when troubleshooting)
        // console.log('[JZ-Dynamic-Content] Loading modified sections:', {
        //     mcpUrl: mcpUrl,
        //     sitePath: sitePath,
        //     filePath: filePath,
        //     fileName: fileName,
        //     requestBody: requestBody
        // });

        fetch(mcpUrl + '/api/read-json-sharepoint', fetchOptions)
        .then(function(response) {
            if (!response.ok) {
                // 404 is expected when file doesn't exist yet - not an error
                if (response.status === 404) {
                    // console.log('[JZ-Dynamic-Content] File not found (expected for new apps):', fileName);
                    return { success: false, status: 404 };
                }

                // Mark as error for non-404 responses
                sharepointErrorCount[errorKey]++;

                // Log the error with response details for debugging (disabled)
                return response.text().then(function(errorText) {
                    // console.error('[JZ-Dynamic-Content] SharePoint API error:', {
                    //     status: response.status,
                    //     statusText: response.statusText,
                    //     url: mcpUrl + '/api/read-json-sharepoint',
                    //     sitePath: sitePath,
                    //     filePath: filePath,
                    //     errorResponse: errorText
                    // });

                    // If 500 error and we've had multiple failures, mark SharePoint as unavailable
                    if (response.status === 500 && sharepointErrorCount[errorKey] >= 3) {
                        if (!sharepointUnavailable) {

                            sharepointUnavailable = true;
                            sharepointUnavailableSince = Date.now();
                        }
                    }

                    // Return empty result for 500 (server error)
                    return { success: false, status: response.status };
                });
            }

            // Success - reset error counter
            sharepointErrorCount[errorKey] = 0;
            if (sharepointUnavailable) {
                
                sharepointUnavailable = false;
                sharepointUnavailableSince = null;
            }

            return response.json();
        })
        .then(function(readResult) {
            if (readResult.success && readResult.data) {
                
                callback(readResult.data);
            } else {
                // File not found (404) or error - this is normal for new PKs
                if (readResult.status === 404 || readResult.status === 500) {
                    // Only log once per file
                    if (sharepointErrorCount[errorKey] === 1 || !readResult.status) {
                        
                    }
                }
                callback(null);
            }
        })
        .catch(function(error) {
            // Network error or other issue
            sharepointErrorCount[errorKey]++;
            if (sharepointErrorCount[errorKey] === 1) {
                
            }
            callback(null);
        });
    }

    // Save content for a SPECIFIC PK value (one file per PK) - NO MERGE NEEDED!
    function saveModifiedSections(appId, spaceId, connectionId, pkValue, sectionLabel, content, callback) {
        if (!pkValue) {
            callback(false, 'No PK value provided');
            return;
        }

        // Load existing data for this PK (to preserve other sections)
        loadModifiedSections(appId, pkValue, function(existingData) {
            // Create structure if new file
            if (!existingData) {
                existingData = {
                    sections: {},
                    appId: appId,
                    pkValue: pkValue,
                    created: new Date().toISOString()
                };
            }

            if (!existingData.sections) {
                existingData.sections = {};
            }

            // Use the globally fetched user info (name + email if available, otherwise internal ID)
            var modifiedByDisplay = currentUserName !== 'Unknown User' && currentUserEmail !== 'user@qlik.com'
                ? currentUserName + ' (' + currentUserEmail + ')'
                : currentUserEmail;

            // Update just this section
            existingData.sections[sectionLabel] = {
                content: content,
                lastModified: new Date().toISOString(),
                modifiedBy: modifiedByDisplay
            };
            existingData.lastModified = new Date().toISOString();

            // Sanitize pkValue for filename (replace special chars with underscores)
            var safePkValue = pkValue.replace(/[^a-zA-Z0-9-_]/g, '_');
            var fileName = 'qlik-sections-' + appId + '-' + safePkValue + '.json';
            var layout = currentModel && currentModel.layout ? currentModel.layout : {};

            var mcpUrl = (layout.mcpOrchestratorUrl || 'https://gse-mcp.replit.app').replace(/\/+$/, '');
            var sitePath = layout.sharepointSitePath || '/sites/QlikTechnicalEnablement';
            var folderPath = layout.sharepointFolderPath || '/Shared Documents/Qlik200_Topsheet';

            

            var headers = { 'Content-Type': 'application/json' };

            // Build fetch options with SSO credentials
            var fetchOptions = {
                method: 'POST',
                headers: headers,
                credentials: 'include', // SSO mode - always include cookies
                body: JSON.stringify({
                    fileName: fileName,
                    data: existingData,
                    sitePath: sitePath,
                    folderPath: folderPath
                })
            };

            fetch(mcpUrl + '/api/store-json-sharepoint', fetchOptions)
            .then(function(response) {
                if (!response.ok) {
                    return response.text().then(function(text) {
                        
                        throw new Error('Failed to save to SharePoint: ' + response.status);
                    });
                }
                return response.json();
            })
            .then(function(result) {
                callback(true, null);
            })
            .catch(function(error) {
                
                callback(false, error.message);
            });
        });
    }

    // Delete a section from SharePoint (removes it from the JSON file)
    function deleteSectionFromSharePoint(appId, spaceId, connectionId, pkValue, sectionLabel, callback) {
        if (!pkValue) {
            callback(false, 'No PK value provided');
            return;
        }

        // Load existing data for this PK
        loadModifiedSections(appId, pkValue, function(existingData) {
            // If no data exists or section doesn't exist, nothing to delete
            if (!existingData || !existingData.sections || !existingData.sections[sectionLabel]) {
                callback(true, null); // Success - section already doesn't exist
                return;
            }

            // Remove the section
            delete existingData.sections[sectionLabel];
            existingData.lastModified = new Date().toISOString();

            // If no sections remain, we could delete the file, but for simplicity just save empty sections
            // This allows the file structure to remain consistent

            // Sanitize pkValue for filename
            var safePkValue = pkValue.replace(/[^a-zA-Z0-9-_]/g, '_');
            var fileName = 'qlik-sections-' + appId + '-' + safePkValue + '.json';
            var layout = currentModel && currentModel.layout ? currentModel.layout : {};

            var mcpUrl = (layout.mcpOrchestratorUrl || 'https://gse-mcp.replit.app').replace(/\/+$/, '');
            var sitePath = layout.sharepointSitePath || '/sites/QlikTechnicalEnablement';
            var folderPath = layout.sharepointFolderPath || '/Shared Documents/Qlik200_Topsheet';

            var headers = { 'Content-Type': 'application/json' };

            var fetchOptions = {
                method: 'POST',
                headers: headers,
                credentials: 'include',
                body: JSON.stringify({
                    fileName: fileName,
                    data: existingData,
                    sitePath: sitePath,
                    folderPath: folderPath
                })
            };

            fetch(mcpUrl + '/api/store-json-sharepoint', fetchOptions)
            .then(function(response) {
                if (!response.ok) {
                    return response.text().then(function(text) {
                        throw new Error('Failed to delete section from SharePoint: ' + response.status);
                    });
                }
                return response.json();
            })
            .then(function(result) {
                callback(true, null);
            })
            .catch(function(error) {
                callback(false, error.message);
            });
        });
    }

    // Fetch current user info from Qlik Cloud (called once at extension initialization)
    function fetchCurrentUser(layout, extensionSelf) {
        if (MULTI_USER_ENABLED !== 1) {
            return;
        }

        try {
            var app = qlik.currApp();

            // First, get the internal user ID
            app.global.getAuthenticatedUser(function(reply) {
                if (reply && reply.qReturn) {
                    var internalUserId = reply.qReturn;

                    // Now fetch full user profile from Qlik Cloud REST API
                    var tenantUrl = window.location.origin; // e.g., https://tenant.region.qlikcloud.com

                    fetch(tenantUrl + '/api/v1/users/me', {
                        method: 'GET',
                        headers: {
                            'Content-Type': 'application/json'
                        },
                        credentials: 'include' // Include cookies for authentication
                    })
                    .then(function(response) {
                        if (!response.ok) {
                            throw new Error('Failed to fetch user profile: ' + response.status);
                        }
                        return response.json();
                    })
                    .then(function(user) {
                        // Store user information
                        currentUserData = user;
                        currentUserName = user.name || 'Unknown User';
                        currentUserEmail = user.email || internalUserId;

                        // Fetch feature permissions now that we have the user email
                        if ((MULTI_USER_ENABLED === 1 || CLAUDE_ENABLED === 1) &&
                            !userPermissionsLoaded &&
                            layout && layout.mcpOrchestratorUrl) {
                            fetchFeaturePermissions(currentUserEmail, layout.mcpOrchestratorUrl, extensionSelf);
                        }
                    })
                    .catch(function(error) {
                        // Fallback: use internal user ID
                        currentUserEmail = internalUserId;

                        // Try to extract readable name from internal ID
                        var userIdMatch = internalUserId.match(/UserId=([^;]+)/);
                        var userId = userIdMatch ? userIdMatch[1] : internalUserId;

                        // If it's an auth0 hash, show truncated version
                        if (userId.startsWith('auth0|') && userId.length > 20) {
                            currentUserName = userId.substring(0, 14) + '...';
                        } else {
                            currentUserName = userId;
                        }

                        // Fetch feature permissions with fallback user ID
                        if ((MULTI_USER_ENABLED === 1 || CLAUDE_ENABLED === 1) &&
                            !userPermissionsLoaded &&
                            layout && layout.mcpOrchestratorUrl) {
                            fetchFeaturePermissions(currentUserEmail, layout.mcpOrchestratorUrl, extensionSelf);
                        }
                    });
                }
            });
        } catch (error) {

        }
    }

    // Check if user has permission for a specific feature (fail-closed: hidden until confirmed)
    function userHasFeaturePermission(featureKey) {
        // If no feature key specified, allow access
        if (!featureKey) {
            return true;
        }

        // If permissions not loaded yet, deny access (fail-closed security model)
        // This prevents flickering of restricted features during page load
        if (!userPermissionsLoaded) {
            return false;
        }

        // Check if feature is in the allowed features array (case-insensitive)
        var featureKeyLower = featureKey.toLowerCase();
        var hasPermission = userAllowedFeatures.some(function(f) {
            return f.toLowerCase() === featureKeyLower;
        });
        return hasPermission;
    }

    // Alias for properties panel (same behavior as runtime features now)
    function userHasPropertyPermission(featureKey) {
        return userHasFeaturePermission(featureKey);
    }

    // Fetch user feature permissions from MCP Orchestrator
    function fetchFeaturePermissions(email, mcpUrl, extensionSelf) {
        if (!email || email === 'user@qlik.com') {
            return;
        }

        if (!mcpUrl) {
            return;
        }

        // Always fetch all features for JZ-Dynamic-Content-Sections
        var appName = 'JZ-Dynamic-Content-Sections';
        var encodedAppName = encodeURIComponent(appName);
        var url = mcpUrl.replace(/\/+$/, '') + '/api/users/me/apps/' + encodedAppName + '/features';

        // Get user info for cross-tenant authentication
        var qlikTenantUrl = window.location.origin;
        var userEmail = email;

        fetch(url, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
                'X-Qlik-User-Email': userEmail,
                'X-Qlik-Tenant-URL': qlikTenantUrl
            },
            credentials: 'include' // Include cookies for same-origin; custom headers for cross-tenant
        })
        .then(function(response) {
            if (!response.ok) {
                throw new Error('HTTP ' + response.status + ': ' + response.statusText);
            }
            return response.json();
        })
        .then(function(data) {
            // Getting all features for app
            if (data.success && data.features) {
                userAllowedFeatures = data.features
                    .filter(function(f) { return f.isGranted === true; })
                    .map(function(f) { return f.featureKey; });
                userPermissionsLoaded = true;
                // CRITICAL: Trigger re-render to show edit buttons now that permissions are loaded
                if (extensionSelf && extensionSelf.backendApi && extensionSelf.backendApi.model) {
                    setTimeout(function() {
                        extensionSelf.backendApi.model.enigmaModel.emit('changed');
                    }, 100);
                }
            } else {
                userAllowedFeatures = [];
                userPermissionsLoaded = true;
                if (extensionSelf && extensionSelf.backendApi && extensionSelf.backendApi.model) {
                    setTimeout(function() {
                        extensionSelf.backendApi.model.enigmaModel.emit('changed');
                    }, 100);
                }
            }
        })
        .catch(function(error) {
            // Mark as failed - will hide buttons for security (fail-closed)
            userAllowedFeatures = [];
            userPermissionsLoaded = true;
            userPermissionsFailed = true;
            if (extensionSelf && extensionSelf.backendApi && extensionSelf.backendApi.model) {
                setTimeout(function() {
                    extensionSelf.backendApi.model.enigmaModel.emit('changed');
                }, 100);
            }
        });
    }

    // Poll for changes from other users (with exponential backoff)
    function startPolling(instanceId, layout, $element, sectionsData, self) {
        if (MULTI_USER_ENABLED !== 1 || MULTI_USER_ENABLED !== 1) {
            return;
        }

        // Stop any existing polling for this instance
        stopPolling(instanceId);

        // Don't poll if no PK field configured
        if (!layout.pkField) {
            return;
        }

        // Get current PK value
        getCurrentPKValue(layout.pkField, function(pkValue, error, pkToken) {
            if (error || !pkValue) {
                return;
            }

            // Initialize polling state
            var baseIntervalSeconds = layout.pollingInterval || (POLLING_INTERVAL / 1000);
            pollingIntervals[instanceId] = baseIntervalSeconds * 1000; // Start with base interval
            consecutiveNoChanges[instanceId] = 0;

            // Set up polling with dynamic interval
            function scheduleNextPoll() {
                // Use fast polling if modal is open, otherwise use exponential backoff interval
                var intervalToUse = modalIsOpen ? FAST_POLLING_INTERVAL : pollingIntervals[instanceId];

                pollingTimers[instanceId] = setTimeout(function() {
                    checkForUpdates(instanceId, layout, $element, sectionsData, self, pkValue, scheduleNextPoll);
                }, intervalToUse);

                if (modalIsOpen) {
                    
                }
            }

            // Start first poll
            scheduleNextPoll();

            
        });
    }

    // Stop polling for an instance
    function stopPolling(instanceId) {
        if (pollingTimers[instanceId]) {
            clearTimeout(pollingTimers[instanceId]);
            delete pollingTimers[instanceId];
            delete pollingIntervals[instanceId];
            delete consecutiveNoChanges[instanceId];
            
        }
    }

    // Adjust polling interval based on activity (exponential backoff)
    function adjustPollingInterval(instanceId, layout, hasChanges) {
        var baseIntervalSeconds = layout.pollingInterval || (POLLING_INTERVAL / 1000);
        var baseIntervalMs = baseIntervalSeconds * 1000;
        var maxIntervalMs = baseIntervalMs * 8; // Max 8x base interval

        if (hasChanges) {
            // Changes detected - reset to base interval (aggressive polling)
            consecutiveNoChanges[instanceId] = 0;
            pollingIntervals[instanceId] = baseIntervalMs;
            
        } else {
            // No changes - increase interval exponentially
            consecutiveNoChanges[instanceId]++;
            var backoffMultiplier = Math.min(Math.pow(2, consecutiveNoChanges[instanceId] - 1), 8);
            pollingIntervals[instanceId] = Math.min(baseIntervalMs * backoffMultiplier, maxIntervalMs);

            if (consecutiveNoChanges[instanceId] <= 3) {
                
            }
        }
    }

    // Check for updates from other users
    function checkForUpdates(instanceId, layout, $element, sectionsData, self, pkValue, scheduleNextPoll) {
        // If SharePoint is unavailable, skip polling to avoid spam
        if (sharepointUnavailable) {
            // Still schedule next poll (will retry after 10 minutes automatically)
            scheduleNextPoll();
            return;
        }

        var app = qlik.currApp();
        var appId = app.model.layout.qFileName.replace('.qvf', '');

        // Load PK-specific file from SharePoint (much faster than loading all records!)
        loadModifiedSections(appId, pkValue, function(result) {
            if (result) {
                var currentData = result; // Direct data, no need for .records[pkValue]
                var lastKnown = lastKnownData[instanceId + '_' + pkValue];
                var isFirstPoll = !lastKnown;

                

                // On first poll, use cached data as baseline if available (NEW STRUCTURE: direct sections)
                if (isFirstPoll && modifiedSectionsCache.loaded && modifiedSectionsCache.data) {
                    lastKnown = modifiedSectionsCache.data;
                    
                }

                // Check if data has changed
                if (lastKnown) {
                    var hasChanges = false;
                    var changedSections = [];

                    sectionsData.forEach(function(sd) {
                        var sectionLabel = sd.section.label;
                        var currentSection = currentData.sections[sectionLabel];
                        var lastSection = lastKnown.sections[sectionLabel];

                        

                        if (currentSection && lastSection) {
                            // Simple check: if the timestamp changed, the content was updated
                            var timestampChanged = currentSection.lastModified !== lastSection.lastModified;

                            

                            if (timestampChanged) {
                        hasChanges = true;
                        changedSections.push({
                                    label: sectionLabel,
                                    modifiedBy: currentSection.modifiedBy,
                                    lastModified: new Date(currentSection.lastModified).toLocaleString()
                                });
                            }
                        }
                    });

                    // If changes detected, show notification and update cache
                    if (hasChanges) {
                        

                        // === MODAL WARNING: If modal is open and editing the changed section ===
                        if (modalIsOpen && modalSectionLabel && modalPKValue === pkValue) {
                            var editingChangedSection = changedSections.some(function(s) {
                        return s.label === modalSectionLabel;
                            });

                            if (editingChangedSection) {
                                // Show prominent warning banner in modal
                        var changedSection = changedSections.find(function(s) { return s.label === modalSectionLabel; });
                        var warningHtml = '<div style="background: #fff3cd; border: 2px solid #ffc107; border-radius: 4px; padding: 12px; margin-bottom: 12px; font-size: 14px; color: #856404;">' +
                                    '<strong>⚠️ WARNING: Content Changed</strong><br>' +
                                    'This section was modified by <strong>' + changedSection.modifiedBy + '</strong><br>' +
                                    'at ' + changedSection.lastModified + '<br><br>' +
                                    '<strong>Your changes will be REJECTED if you try to save.</strong><br>' +
                                    'Please copy your changes and close this modal to see the latest version.' +
                                    '</div>';

                                // Find modal body and prepend warning (remove old warning first)
                        var $modalBody = $('.edit-modal .modal-body');
                                $modalBody.find('.conflict-warning').remove();
                                $modalBody.prepend('<div class="conflict-warning">' + warningHtml + '</div>');

                        
                            }
                        }

                        showUpdateNotification($element, changedSections);

                        // Update cache and trigger re-render (NEW STRUCTURE: result is direct data)
                        modifiedSectionsCache.data = result;
                        modifiedSectionsCache.loaded = true;
                        modifiedSectionsCache.pkValue = pkValue;

                        

                        // Re-render to show updated content
                        renderSectionsUI($element, layout, sectionsData, self, instanceId);

                        
                    }

                    // Adjust polling interval based on whether changes were detected
                    adjustPollingInterval(instanceId, layout, hasChanges);
                } else {
                    // First poll - log what we found
                    
                    Object.keys(currentData.sections || {}).forEach(function(sectionLabel) {
                        var section = currentData.sections[sectionLabel];
                        
                    });

                    // First poll - no previous data to compare
                    adjustPollingInterval(instanceId, layout, false);
                }

                // Store current data as last known
                lastKnownData[instanceId + '_' + pkValue] = JSON.parse(JSON.stringify(currentData));
            } else {
                // No file exists yet for this PK (first time editing)
                
                adjustPollingInterval(instanceId, layout, false);
            }

            // Schedule next poll
            if (scheduleNextPoll) {
                scheduleNextPoll();
            }
        });
    }

    // Show notification when content is updated by another user
    function showUpdateNotification($element, changedSections) {
        var $notification = $('<div class="update-notification"></div>');

        var message = '<strong>🔄 Content Updated</strong><br>';
        changedSections.forEach(function(section) {
            message += '<small>' + section.label + ' modified by ' + section.modifiedBy +
                      ' at ' + section.lastModified + '</small><br>';
        });

        $notification.html(message);
        $notification.css({
            'position': 'fixed',
            'top': '20px',
            'right': '20px',
            'background': '#e3f2fd',
            'border-left': '4px solid #1976d2',
            'padding': '12px 16px',
            'border-radius': '4px',
            'box-shadow': '0 2px 8px rgba(0,0,0,0.15)',
            'z-index': '10001',
            'max-width': '350px',
            'animation': 'slideInRight 0.3s ease'
        });

        $('body').append($notification);

        // Auto-hide after 8 seconds
        setTimeout(function() {
            $notification.fadeOut(300, function() {
                $notification.remove();
            });
        }, 8000);
    }

    function openSectionEditModal(section, sectionData, layout, self, renderedContent, originalTemplateHtml, clickPaintCounter, $element, sectionsData, instanceId) {
        // console.log('[MODAL] openSectionEditModal called for section:', section ? section.label : 'NULL');
        // console.log('[MODAL] Parameters:', {
        //     hasSection: !!section,
        //     hasSectionData: !!sectionData,
        //     hasLayout: !!layout,
        //     hasSelf: !!self,
        //     pkField: layout ? layout.pkField : 'N/A',
        //     clickPaintCounter: clickPaintCounter,
        //     globalPaintCounter: globalPaintCounter
        // });

        // Capture stack trace to see WHO is calling this function
        var caller = 'unknown';
        try {
            var stack = new Error().stack;
            if (stack) {
                var lines = stack.split('\n');
                caller = lines[2] ? lines[2].trim() : 'unknown';
            }
        } catch(e) {
            caller = 'error_getting_stack';
        }

        var app = qlik.currApp();

        // Check if modal is already open
        if ($('.edit-modal-overlay').length > 0) {
            return;
        }

        // Check if PK field is configured
        if (!layout.pkField) {
            alert('⚠️ Primary Key Field not configured.\n\nPlease set the "Primary Key Field" in API Settings.');
            return;
        }

        // Get PK value from current selections
        var pkCheckResult = getCurrentPKValue(layout.pkField, function(pkValue, error, pkToken) {
            // CRITICAL: Check if this token already has an open modal
            if (pkToken === currentModalPKToken) {
                return;
            }

            // CRITICAL: Check if paint() was called after button click
            // If globalPaintCounter changed, a new paint happened (selection changed)
            if (clickPaintCounter !== undefined && globalPaintCounter !== clickPaintCounter) {
                return;
            }

            if (error) {
                return;
            }

            if (!pkValue) {
                alert('⚠️ No Primary Key selected.\n\nPlease select exactly ONE value in the "' + layout.pkField + '" field to edit this section.');
                return;
            }

            // === CRITICAL FIX: Always load fresh data and UPDATE VIEW before opening modal ===
            // This ensures the extension view shows the latest content BEFORE the modal opens
            

            var appId = app.id;
            loadModifiedSections(appId, pkValue, function(freshData) {
                var contentToEdit = renderedContent || ''; // Fallback to rendered content
                var freshTimestamp = null;
                var needsRerender = false;
                var hasSavedContent = false; // Track if this section has been edited before

                // Check if there's saved content in the fresh data (NEW STRUCTURE: direct sections)
                if (freshData &&
                    freshData.sections &&
                    freshData.sections[section.label]) {

                    contentToEdit = freshData.sections[section.label].content;
                    freshTimestamp = freshData.sections[section.label].lastModified;
                    hasSavedContent = true; // Mark that saved content exists

                    

                    // Check if cache needs updating (content changed since last load)
                    if (!modifiedSectionsCache.loaded ||
                        !modifiedSectionsCache.data ||
                        !modifiedSectionsCache.data.sections ||
                        !modifiedSectionsCache.data.sections[section.label] ||
                        modifiedSectionsCache.data.sections[section.label].lastModified !== freshTimestamp) {

                        needsRerender = true;
                        
                    }

                    // Update cache with fresh data
                    modifiedSectionsCache.data = freshData;
                    modifiedSectionsCache.loaded = true;
                    modifiedSectionsCache.pkValue = pkValue;
                } else {

                }

                // Clean up contentToEdit by removing [No data] placeholders, horizontal rules, and label separators
                contentToEdit = contentToEdit.replace(/\[No [Dd]ata\]/gi, '');
                contentToEdit = contentToEdit.replace(/<span[^>]*>\[No [Dd]ata\]<\/span>/gi, '');
                contentToEdit = contentToEdit.replace(/<hr[^>]*>/gi, '');
                // Remove separator divs (height: 2px gradient lines)
                contentToEdit = contentToEdit.replace(/<div[^>]*height:\s*2px[^>]*><\/div>/gi, '');

                // === STEP 1: Update the extension view FIRST if content changed ===
                if (needsRerender) {
                    

                    // Re-render the entire extension to show fresh content
                    renderSectionsUI($element, layout, sectionsData, self, instanceId);

                    

                    // Small delay to let DOM update before opening modal
                    setTimeout(function() {
                        openModalWithFreshContent();
                    }, 100);
                } else {
                    // No update needed, open modal immediately
                    openModalWithFreshContent();
                }

                function openModalWithFreshContent() {

                    // === STEP 2: Now open modal with fresh content ===
                    // === CONFLICT DETECTION: Capture baseline timestamp ===
                    // Use the fresh timestamp we just loaded
                    modalBaselineTimestamp = freshTimestamp;
                    modalSectionLabel = section.label;
                    modalPKValue = pkValue;

                    if (modalBaselineTimestamp) {
                        
                    } else {
                        
                    }

                    // === STRIP LABEL from editor content if showLabel is enabled ===
                    // The label should be displayed separately (non-editable), not included in editable content
                    var contentToEditStripped = contentToEdit;
                    if (section.showLabel && section.label) {
                        // Build the label HTML pattern to remove (same as rendering logic)
                        // Calculate label size: default font + offset
                        var baseFontSize = parseInt(layout.fontSize || '14', 10);
                        var labelSizeOffset = parseInt(layout.labelSizeOffset || '4', 10);
                        var labelSize = (baseFontSize + labelSizeOffset) + 'px';
                        var labelColor = section.labelColor || '#1a1a1a';
                        var labelSeparator = (section.labelSeparator === true || section.labelSeparator === 'true');

                        // Remove label div and optional separator from the beginning
                        var labelPattern = '<div style="font-size: ' + labelSize + '; color: ' + labelColor;
                        var labelEndPattern = section.label + '</div>';

                        if (contentToEditStripped.indexOf(labelPattern) === 0 || contentToEditStripped.indexOf('<div style="font-size:') === 0) {
                            // Find the end of the label div
                            var labelEnd = contentToEditStripped.indexOf(labelEndPattern);
                            if (labelEnd > -1) {
                        contentToEditStripped = contentToEditStripped.substring(labelEnd + labelEndPattern.length);

                                // Also remove separator if present (check for both 1px and 2px heights)
                        if (labelSeparator && (contentToEditStripped.indexOf('<div style="height: 2px;') === 0 || contentToEditStripped.indexOf('<div style="height: 1px;') === 0)) {
                                    var separatorEnd = contentToEditStripped.indexOf('</div>');
                                    if (separatorEnd > -1) {
                                        contentToEditStripped = contentToEditStripped.substring(separatorEnd + 6);
                                    }
                                }
                            }
                        }

                        
                    }

                    // === FAST POLLING: Enable when modal opens ===
                    modalIsOpen = true;
                    

                    // Create modal overlay
                    var $overlay = $('<div class="edit-modal-overlay"></div>');
                    var $modal = $('<div class="edit-modal"></div>');

                    var $header = $('<div class="modal-header"></div>');
                    $header.append('<h3>Edit Section: ' + (section.label || 'Untitled') + '</h3>');
                    var $closeBtn = $('<button class="modal-close-btn">&times;</button>');
                    $closeBtn.on('click', function() {
                        closeModalAndCleanup();
                    });
                    $header.append($closeBtn);

                    var $body = $('<div class="modal-body"></div>');

                    // Create formatted editor showing rendered content
                    var $editorContainer = $('<div style="display: flex; flex-direction: column; height: 100%;"></div>');

                    // === DISPLAY LABEL SEPARATELY (non-editable) if showLabel is enabled ===
                    if (section.showLabel && section.label) {
                        var labelStyle = section.labelStyle || 'bold';
                        // Calculate label size: default font + offset
                        var baseFontSize = parseInt(layout.fontSize || '14', 10);
                        var labelSizeOffset = parseInt(layout.labelSizeOffset || '4', 10);
                        var labelSize = (baseFontSize + labelSizeOffset) + 'px';
                        var labelColor = section.labelColor || '#1a1a1a';
                        var labelSeparator = (section.labelSeparator === true || section.labelSeparator === 'true');

                        var styleAttr = 'font-size: ' + labelSize + '; color: ' + labelColor + '; margin: 0; line-height: 1.3; padding: 0 0 1px 0; background: #f9f9f9; padding: 8px; border-radius: 4px;';
                        if (labelStyle === 'bold') {
                            styleAttr += ' font-weight: 600;';
                        } else if (labelStyle === 'italic') {
                            styleAttr += ' font-style: italic;';
                        }

                        var $labelDisplay = $('<div style="' + styleAttr + '">' + section.label + '</div>');
                        $editorContainer.append($labelDisplay);
                    }

                    // === CREATE EDITOR FIRST (so toolbar buttons can reference it) ===
                    // Apply default font, size, and color from appearance settings
                    var defaultFont = layout.fontFamily || 'Arial, sans-serif';
                    var defaultSizeNum = layout.fontSize || '14';
                    var defaultSize = defaultSizeNum.toString().indexOf('px') > -1 ? defaultSizeNum : defaultSizeNum + 'px';
                    var defaultColor = '#333333'; // Dark grey to match component default
                    var $editor = $('<div class="markdown-content" contenteditable="true" style="flex: 1; padding: 45px 12px 12px 12px; border: 2px solid #d0d0d0; border-radius: 4px; background: white; overflow-y: auto; min-height: 400px; outline: none; transition: border-color 0.2s ease; caret-color: #1976d2; line-height: 1.8; font-family: ' + defaultFont + '; font-size: ' + defaultSize + '; color: ' + defaultColor + ';"></div>');

                    // Add CSS for better cursor visibility
                    var $editorCSS = $('<style>' +
                        '.markdown-content:focus { border-color: #1976d2 !important; }' +
                        '.markdown-content .master-item-badge, ' +
                        '.markdown-content .tag-badge { margin: 0 4px; }' +
                        '</style>');
                    if ($('head style:contains("markdown-content")').length === 0) {
                        $('head').append($editorCSS);
                    }

                    // Add dynamic CSS for list styling with current font settings
                    var editorId = 'editor-' + Date.now();
                    $editor.attr('id', editorId);
                    var $listCSS = $('<style id="list-style-' + editorId + '">' +
                        '#' + editorId + ' ul, #' + editorId + ' ol { ' +
                        '  font-family: ' + defaultFont + '; ' +
                        '  font-size: ' + defaultSize + '; ' +
                        '  color: ' + defaultColor + '; ' +
                        '  line-height: 1.8; ' +
                        '  margin: 0.5em 0; ' +
                        '  padding-left: 28px; ' +
                        '}' +
                        '#' + editorId + ' ul li, #' + editorId + ' ol li { ' +
                        '  font-family: inherit; ' +
                        '  font-size: inherit; ' +
                        '  color: inherit; ' +
                        '}' +
                        '</style>');
                    $('head').append($listCSS);

                    // === FORMATTING TOOLBAR ===
                    var $toolbar = $('<div class="editor-toolbar"></div>');

                    // Helper function to create toolbar button
                    function createToolbarBtn(label, command, value) {
                        var $btn = $('<button type="button"></button>');
                        $btn.html(label);
                        $btn.on('mousedown', function(e) {
                            e.preventDefault(); // Prevent losing focus from editor
                        });
                        $btn.on('click', function(e) {
                            e.preventDefault();
                            document.execCommand(command, false, value);
                            $editor.focus(); // Return focus to editor
                        });
                        return $btn;
                    }

                    // Text formatting buttons
                    $toolbar.append(createToolbarBtn('<b>B</b>', 'bold'));
                    $toolbar.append(createToolbarBtn('<i>I</i>', 'italic'));
                    $toolbar.append(createToolbarBtn('<u>U</u>', 'underline'));

                    // Separator
                    $toolbar.append('<div class="toolbar-separator"></div>');

                    // Alignment buttons
                    $toolbar.append(createToolbarBtn('⬅️', 'justifyLeft'));
                    $toolbar.append(createToolbarBtn('↔️', 'justifyCenter'));
                    $toolbar.append(createToolbarBtn('➡️', 'justifyRight'));

                    // Separator
                    $toolbar.append('<div class="toolbar-separator"></div>');

                    // Font size dropdown with pixel values (matches extension settings)
                    var $fontSizeSelect = $('<select title="Font Size"></select>');
                    // Show current default size in placeholder
                    var currentSizeLabel = defaultSize.indexOf('px') > -1 ? defaultSize : defaultSize + 'px';
                    $fontSizeSelect.append('<option value="">Font Size (default: ' + currentSizeLabel + ')</option>');
                    var sizes = ['10', '12', '14', '16', '18', '20', '22', '24', '28', '32', '36', '40', '48', '56', '64', '72'];
                    sizes.forEach(function(size) {
                        $fontSizeSelect.append('<option value="' + size + 'px">' + size + 'px</option>');
                    });
                    $fontSizeSelect.on('change', function() {
                        var size = $(this).val();
                        if (size) {
                            var selection = window.getSelection();

                            // Check if all content is selected (CTRL+A scenario)
                            if (selection.rangeCount > 0) {
                                var range = selection.getRangeAt(0);
                                var editorEl = $editor[0];

                                // Check if selection spans the entire editor
                                var isFullSelection = false;
                                try {
                                    // Selection is "full" if it starts at the beginning and ends at the end
                                    var isAtStart = range.startOffset === 0 && range.startContainer === editorEl;
                                    var isAtEnd = range.endOffset === range.endContainer.length &&
                                                  (range.endContainer === editorEl ||
                                                   range.endContainer.parentNode === editorEl);

                                    // Also check if the selection text length is close to editor text length
                                    var selectedText = range.toString().trim();
                                    var allText = $editor.text().trim();

                                    isFullSelection = (selectedText.length >= allText.length * 0.9) ||
                                                     (isAtStart && isAtEnd);
                                } catch(e) {
                                    // Fallback to text comparison
                                    var selectedText = range.toString().trim();
                                    var allText = $editor.text().trim();
                                    isFullSelection = selectedText.length >= allText.length * 0.9;
                                }

                                if (isFullSelection) {
                                    // Apply font size to the entire editor and all its children
                                    // Preserve default color when changing size
                                    $editor.css('font-size', size);
                                    $editor.find('*').each(function() {
                                        $(this).css('font-size', size);
                                        // Only apply default color if element doesn't have explicit color
                                        if (!$(this).attr('style') || !$(this).attr('style').includes('color:')) {
                                            $(this).css('color', defaultColor);
                                        }
                                    });
                                } else if (!selection.isCollapsed) {
                                    // Normal selection - wrap in span
                                    try {
                                        var span = document.createElement('span');
                                        span.style.fontSize = size;
                                        range.surroundContents(span);
                                    } catch(e) {
                                        // Fallback if surroundContents fails (complex selection)
                                        document.execCommand('fontSize', false, '7');
                                        var fontElements = $editor.find('font[size="7"]');
                                        fontElements.removeAttr('size').css('font-size', size);
                                    }
                                }
                            }
                            $(this).val(''); // Reset dropdown
                            $editor.focus();
                        }
                    });
                    $toolbar.append($fontSizeSelect);

                    // Text color picker
                    var savedSelection = null;
                    var $colorPickerWrapper = $('<div style="display: inline-flex; align-items: center; gap: 4px; padding: 0 4px;"></div>');
                    var $colorPicker = $('<input type="color" value="' + defaultColor + '" title="Text Color (default: ' + defaultColor + ')" style="cursor: pointer;">');

                    // Save selection when clicking color picker
                    $colorPicker.on('mousedown', function(e) {
                        // Save current selection before color picker opens
                        var selection = window.getSelection();
                        if (selection.rangeCount > 0) {
                            savedSelection = selection.getRangeAt(0);
                        }
                    });

                    // Apply color when changed
                    $colorPicker.on('input change', function() {
                        var color = $(this).val();

                        // Restore selection
                        if (savedSelection) {
                            $editor.focus();
                            var selection = window.getSelection();
                            selection.removeAllRanges();
                            selection.addRange(savedSelection);
                        }

                        // Apply color
                        document.execCommand('foreColor', false, color);
                        
                    });

                    $colorPickerWrapper.append($colorPicker);
                    $colorPickerWrapper.append('<span style="font-size: 11px; color: #666;">A</span>');
                    $toolbar.append($colorPickerWrapper);

                    // Background/Highlight color picker
                    var savedBgSelection = null;
                    var $bgColorPickerWrapper = $('<div style="display: inline-flex; align-items: center; gap: 4px; padding: 0 4px;"></div>');
                    var $bgColorPicker = $('<input type="color" value="#ffff00" title="Background/Highlight Color" style="cursor: pointer;">');

                    // Save selection when clicking bg color picker
                    $bgColorPicker.on('mousedown', function(e) {
                        // Save current selection before color picker opens
                        var selection = window.getSelection();
                        if (selection.rangeCount > 0) {
                            savedBgSelection = selection.getRangeAt(0);
                        }
                    });

                    // Apply background color when changed
                    $bgColorPicker.on('input change', function() {
                        var color = $(this).val();

                        // Restore selection
                        if (savedBgSelection) {
                            $editor.focus();
                            var selection = window.getSelection();
                            selection.removeAllRanges();
                            selection.addRange(savedBgSelection);
                        }

                        // Apply background color
                        document.execCommand('backColor', false, color);
                        
                    });

                    $bgColorPickerWrapper.append($bgColorPicker);
                    $bgColorPickerWrapper.append('<span style="font-size: 11px; color: #666; background: #ffff00; padding: 0 2px;">A</span>');
                    $toolbar.append($bgColorPickerWrapper);

                    // Separator
                    $toolbar.append('<div class="toolbar-separator"></div>');

                    // List buttons
                    $toolbar.append(createToolbarBtn('• List', 'insertUnorderedList'));
                    $toolbar.append(createToolbarBtn('1. List', 'insertOrderedList'));

                    // Separator
                    $toolbar.append('<div class="toolbar-separator"></div>');

                    // Clear formatting button
                    var $clearBtn = createToolbarBtn('✖ Clear', 'removeFormat');
                    $clearBtn.attr('title', 'Remove all formatting from selected text');
                    $toolbar.append($clearBtn);

                    // Separator
                    $toolbar.append('<div class="toolbar-separator"></div>');

                    // Undo button
                    var $undoBtn = $('<button type="button" title="Undo last change"></button>');
                    $undoBtn.html('← Undo');
                    $undoBtn.on('mousedown', function(e) {
                        e.preventDefault();
                    });
                    $undoBtn.on('click', function(e) {
                        e.preventDefault();
                        document.execCommand('undo');
                        $editor.focus();
                    });
                    $toolbar.append($undoBtn);

                    // Reset to original template button
                    var $resetBtn = $('<button type="button" title="Reset to original template with live data values" style="background: #fff3cd; border-color: #ffc107;"></button>');
                    $resetBtn.html('↻ Reset');
                    $resetBtn.on('mousedown', function(e) {
                        e.preventDefault();
                    });
                    $resetBtn.on('click', function(e) {
                        e.preventDefault();
                        if (confirm('Reset to original template?\n\nThis will delete your saved edits and restore the original Qlik template with live data values.\n\nThis action cannot be undone.')) {
                            // Disable buttons during reset
                            $resetBtn.prop('disabled', true);
                            $saveBtn.prop('disabled', true);
                            $statusMsg.text('Resetting...').removeClass('error success').css('color', '#666');

                            // Set flag to prevent paint() during reset
                            savingInProgress = true;

                            // Get appId
                            var app = qlik.currApp();
                            var appId = app.id;

                            // Delete the section from SharePoint
                            deleteSectionFromSharePoint(appId, null, null, pkValue, section.label, function(success, errorMsg) {
                                if (success) {
                                    $statusMsg.text('✅ Reset successful!').addClass('success');

                                    // Reload fresh data from SharePoint to update cache
                                    // The section was already deleted in SharePoint, so fresh data will reflect that
                                    loadModifiedSections(appId, pkValue, function(freshData) {
                                        // Update cache with fresh data (without the deleted section)
                                        // The delete operation in SharePoint should have left other sections intact
                                        if (freshData && freshData.sections) {
                                            modifiedSectionsCache.data = freshData;
                                            modifiedSectionsCache.loaded = true;
                                            modifiedSectionsCache.pkValue = pkValue;
                                        } else {
                                            // No fresh data or no sections - clear cache entirely
                                            modifiedSectionsCache.data = null;
                                            modifiedSectionsCache.loaded = false;
                                            modifiedSectionsCache.pkValue = null;
                                        }

                                        // Force a full re-render to show original template with fresh Qlik data
                                        renderSectionsUI($element, layout, sectionsData, self, instanceId);

                                        // Close modal after brief delay
                                        setTimeout(function() {
                                            closeModalAndCleanup();
                                            savingInProgress = false;
                                        }, 800);
                                    }); // End loadModifiedSections callback
                                } else {
                                    $statusMsg.text('❌ Reset failed: ' + errorMsg).addClass('error');
                                    $resetBtn.prop('disabled', false);
                                    $saveBtn.prop('disabled', false);
                                    savingInProgress = false;
                                }
                            });
                        }
                    });
                    $resetBtn.hover(
                        function() { $(this).css({'background': '#ffe082', 'border-color': '#ffa000'}); },
                        function() { $(this).css({'background': '#fff3cd', 'border-color': '#ffc107'}); }
                    );
                    $toolbar.append($resetBtn);

                    $editorContainer.append($toolbar);

                    // Add focus/blur styling to editor
                $editor.on('focus', function() {
                    $(this).css('border-color', '#009845');
                }).on('blur', function() {
                    $(this).css('border-color', '#d0d0d0');
                });

                // === SINGLE EDITOR with Protected Qlik Data Values ===
                // Get the original markdown (the actual section content template)
                // console.log('[EDIT MODAL DEBUG] section object:', section);
                // console.log('[EDIT MODAL DEBUG] section.markdownText:', section.markdownText);
                // console.log('[EDIT MODAL DEBUG] contentToEditStripped:', contentToEditStripped);
                var originalMarkdown = section.markdownText || '';

                // Get the rendered content and mark Qlik data values as protected
                var editorContent = contentToEditStripped;
                if (editorContent === '<span class="no-data">No data</span>' ||
                    editorContent.trim() === 'No data' ||
                    editorContent.indexOf('class="no-data"') > -1) {
                    editorContent = '';
                }
                editorContent = editorContent.replace(/\[No [Dd]ata\]/gi, '');
                editorContent = editorContent.replace(/<span[^>]*>\[No [Dd]ata\]<\/span>/gi, '');
                // Remove horizontal rules (---) and label separators from editor
                editorContent = editorContent.replace(/<hr[^>]*>/gi, '');
                // Remove separator divs (height: 2px gradient lines)
                editorContent = editorContent.replace(/<div[^>]*height:\s*2px[^>]*><\/div>/gi, '');
                editorContent = editorContent.trim();

                // Wrap section content in a deletable blue badge
                if (originalMarkdown && originalMarkdown.indexOf('{{[') > -1) {
                    // Check if content already has the placeholder (subsequent edits)
                    if (editorContent.indexOf('⟨⟨SECTION_CONTENT⟩⟩') > -1) {
                        // Re-render the template with fresh data to show actual content
                        var freshTemplateContent = originalMarkdown;

                        // Get the data from sectionData parameter
                        var data = sectionData ? sectionData.data : null;
                        var numDimensions = sectionData ? sectionData.numDimensions : 0;
                        var numMeasures = sectionData ? sectionData.numMeasures : 0;
                        var itemMapping = sectionData ? sectionData.itemMapping : {};

                        // STEP 1: Convert {{[Master Item Name]}} to {{dim1}}, {{measure1}} format
                        Object.keys(itemMapping).forEach(function(itemName) {
                            var mapping = itemMapping[itemName];
                            var masterPattern = '{{[' + itemName + ']}}';
                            var indexPattern;

                            if (mapping.type === 'dim') {
                                indexPattern = '{{dim' + (mapping.index + 1) + '}}';
                            } else {
                                indexPattern = '{{measure' + (mapping.index + 1) + '}}';
                            }

                            var regex = new RegExp(masterPattern.replace(/[{}[\]]/g, '\\$&'), 'g');
                            freshTemplateContent = freshTemplateContent.replace(regex, indexPattern);
                        });

                        // STEP 2: Replace placeholders with actual values (outside iterating tags)
                        if (data && data.qMatrix && data.qMatrix.length > 0) {
                            function isInsideIteratingTagEditor(text, matchIndex) {
                                var iteratingTags = ['list', 'table', 'concat', 'grid', 'kpi'];
                                var before = text.substring(0, matchIndex);
                                var lastOpenMatch = null;
                                var lastOpenIndex = -1;
                                iteratingTags.forEach(function(tag) {
                                    var regex = new RegExp('#\\[' + tag + '[^\\]]*\\]', 'g');
                                    var match;
                                    while ((match = regex.exec(before)) !== null) {
                                        if (match.index > lastOpenIndex) {
                                            lastOpenIndex = match.index;
                                            lastOpenMatch = tag;
                                        }
                                    }
                                });
                                if (lastOpenMatch) {
                                    var afterOpen = before.substring(lastOpenIndex);
                                    var closeRegex = new RegExp('#\\[\\/' + lastOpenMatch + '\\]');
                                    if (!closeRegex.test(afterOpen)) {
                                        return true;
                                    }
                                }
                                return false;
                            }

                            for (var dimIdx = 0; dimIdx < numDimensions; dimIdx++) {
                                var placeholder = '{{dim' + (dimIdx + 1) + '}}';
                                var regex = new RegExp(placeholder.replace(/[{}]/g, '\\$&'), 'g');
                                var cell = data.qMatrix[0][dimIdx];
                                var value = cell ? cell.qText : 'N/A';
                                freshTemplateContent = freshTemplateContent.replace(regex, function(match, offset) {
                                    return isInsideIteratingTagEditor(freshTemplateContent, offset) ? match : value;
                                });
                            }

                            for (var meaIdx = 0; meaIdx < numMeasures; meaIdx++) {
                                var placeholder = '{{measure' + (meaIdx + 1) + '}}';
                                var regex = new RegExp(placeholder.replace(/[{}]/g, '\\$&'), 'g');
                                var cell = data.qMatrix[0][numDimensions + meaIdx];
                                var value = cell ? cell.qText : 'N/A';
                                freshTemplateContent = freshTemplateContent.replace(regex, function(match, offset) {
                                    return isInsideIteratingTagEditor(freshTemplateContent, offset) ? match : value;
                                });
                            }
                        } else {
                            // No data - replace with [No data] message
                            freshTemplateContent = freshTemplateContent.replace(/\{\{dim\d+\}\}/g, '<span style="color: #999; font-style: italic;">[No data]</span>');
                            freshTemplateContent = freshTemplateContent.replace(/\{\{measure\d+\}\}/g, '<span style="color: #999; font-style: italic;">[No data]</span>');
                            freshTemplateContent = freshTemplateContent.replace(/\{\{\[([^\]]+)\]\}\}/g, '<span style="color: #f57c00; font-style: italic;" title="Master item not found">[Master item &quot;$1&quot; not found]</span>');
                        }

                        // STEP 3: Process content tags in the template
                        if (freshTemplateContent.indexOf('#[') !== -1) {
                            freshTemplateContent = processContentTags(freshTemplateContent, data, numDimensions, numMeasures, layout, null, itemMapping);
                        }

                        // STEP 4: Parse the template markdown
                        var freshRendered = parseMarkdown(freshTemplateContent);

                        // STEP 5: Remove horizontal rules and label separators from rendered content
                        freshRendered = freshRendered.replace(/<hr[^>]*>/gi, '');
                        // Remove separator divs (height: 2px gradient lines)
                        freshRendered = freshRendered.replace(/<div[^>]*(?:min-)?height:\s*2px[^>]*><\/div>/gi, '');

                        // STEP 6: Replace placeholder with blue badge containing actual rendered content
                        editorContent = editorContent.replace(/⟨⟨SECTION_CONTENT⟩⟩/g,
                            '<span class="qlik-section-content" contenteditable="false" data-original-template="' +
                            originalMarkdown.replace(/"/g, '&quot;') + '" ' +
                            'title="Data from Qlik (deletable - press Delete or Backspace)" ' +
                            'style="background-color: #e3f2fd; padding: 2px 6px; border-radius: 3px; border: 1px solid #90caf9; display: inline-block; cursor: pointer; user-select: all;">' +
                            freshRendered + '</span>');
                    } else if (!hasSavedContent) {
                        // First time edit ONLY (no saved content exists) - wrap entire rendered content in badge
                        var $tempDiv = $('<div>').html(editorContent);
                        var textContent = $tempDiv.text().trim();

                        if (textContent) {
                            editorContent = '<span class="qlik-section-content" contenteditable="false" data-original-template="' +
                                           originalMarkdown.replace(/"/g, '&quot;') + '" ' +
                                           'title="Data from Qlik (deletable - press Delete or Backspace)" ' +
                                           'style="background-color: #e3f2fd; padding: 2px 6px; border-radius: 3px; border: 1px solid #90caf9; display: inline-block; cursor: pointer; user-select: all;">' +
                                           editorContent + '</span>';
                        }
                    }
                    // else: hasSavedContent is true but no placeholder means user converted/deleted the badge
                    // In this case, don't wrap anything - just show the saved editable content as-is
                }

                $editor.html(editorContent);

                // Add action buttons to all Qlik badges
                $editor.find('.qlik-section-content').each(function() {
                    var $badge = $(this);

                    // Add buttons directly to the badge
                    var $actions = $('<span class="badge-actions"></span>');
                    $actions.css({
                        'position': 'absolute',
                        'top': '-30px',
                        'right': '0',
                        'background': 'white',
                        'border': '1px solid #90caf9',
                        'border-radius': '4px',
                        'padding': '4px',
                        'display': 'flex',
                        'gap': '4px',
                        'box-shadow': '0 2px 8px rgba(0,0,0,0.15)',
                        'z-index': '10000'
                    });

                    var $convertBtn = $('<button title="Convert to editable text (removes Qlik updates)">🔓 Convert</button>');
                    $convertBtn.css({
                        'padding': '4px 8px',
                        'background': '#4caf50',
                        'color': 'white',
                        'border': 'none',
                        'border-radius': '3px',
                        'cursor': 'pointer',
                        'font-size': '11px',
                        'font-weight': 'bold',
                        'white-space': 'nowrap'
                    });

                    var $deleteBtn = $('<button title="Delete this content">🗑️ Delete</button>');
                    $deleteBtn.css({
                        'padding': '4px 8px',
                        'background': '#f44336',
                        'color': 'white',
                        'border': 'none',
                        'border-radius': '3px',
                        'cursor': 'pointer',
                        'font-size': '11px',
                        'font-weight': 'bold',
                        'white-space': 'nowrap'
                    });

                    $convertBtn.on('click', function(e) {
                        e.stopPropagation();
                        e.preventDefault();

                        // Extract the inner content (exclude the action buttons)
                        var $clone = $badge.clone();
                        $clone.find('.badge-actions').remove();
                        var innerContent = $clone.html();

                        // Replace the badge with its content (making it editable)
                        $badge.replaceWith(innerContent);

                        $editor.focus();
                    });

                    $deleteBtn.on('click', function(e) {
                        e.stopPropagation();
                        e.preventDefault();
                        $badge.remove();
                        $editor.focus();
                    });

                    $actions.append($convertBtn);
                    $actions.append($deleteBtn);

                    $badge.css('position', 'relative');
                    $badge.append($actions);
                });

                $editorContainer.append($editor);
                $body.append($editorContainer);

                // Handle Delete and Backspace keys to remove the qlik-section-content badge and empty divs
                $editor.on('keydown', function(e) {
                    // Check if Delete (46) or Backspace (8) was pressed
                    if (e.keyCode === 46 || e.keyCode === 8 || e.key === 'Delete' || e.key === 'Backspace') {
                        var selection = window.getSelection();
                        if (selection.rangeCount > 0) {
                            var range = selection.getRangeAt(0);
                            var selectedNode = range.commonAncestorContainer;

                            // Find the badge element if we're inside it or have it selected
                            var $badge = null;
                            if (selectedNode.nodeType === 3) { // Text node
                                $badge = $(selectedNode).closest('.qlik-section-content');
                            } else {
                                $badge = $(selectedNode).hasClass('qlik-section-content') ? $(selectedNode) : $(selectedNode).find('.qlik-section-content');
                            }

                            // Also check if the entire badge is selected (user-select: all)
                            if (!$badge || $badge.length === 0) {
                                var $container = $(range.commonAncestorContainer);
                                if ($container.hasClass('qlik-section-content')) {
                                    $badge = $container;
                                } else {
                                    $badge = $container.find('.qlik-section-content').filter(function() {
                                        return selection.containsNode(this, false);
                                    });
                                }
                            }

                            if ($badge && $badge.length > 0) {
                                e.preventDefault();
                                $badge.remove();
                                return false;
                            }

                            // Check if cursor is in an empty div with just <br>
                            var $currentElement = $(selectedNode.nodeType === 3 ? selectedNode.parentNode : selectedNode);
                            if ($currentElement.is('div')) {
                                var html = $currentElement.html();
                                if (html === '<br>' || html === '<br/>' || html.trim() === '') {
                                    e.preventDefault();
                                    // Move cursor to previous element if it exists
                                    var $prev = $currentElement.prev();
                                    $currentElement.remove();
                                    if ($prev.length > 0) {
                                        // Set cursor at end of previous element
                                        var prevNode = $prev[0];
                                        var newRange = document.createRange();
                                        newRange.selectNodeContents(prevNode);
                                        newRange.collapse(false);
                                        selection.removeAllRanges();
                                        selection.addRange(newRange);
                                    }
                                    return false;
                                }
                            }
                        }
                    }
                });

                // Handle paste events - paste plain text only to preserve current formatting
                $editor.on('paste', function(e) {
                    e.preventDefault();

                    // Get plain text from clipboard
                    var text = '';
                    if (e.originalEvent && e.originalEvent.clipboardData) {
                        text = e.originalEvent.clipboardData.getData('text/plain');
                    } else if (window.clipboardData) {
                        // IE fallback
                        text = window.clipboardData.getData('Text');
                    }

                    // Clean up text: trim and collapse multiple consecutive newlines
                    if (text) {
                        // Replace multiple consecutive newlines (2 or more) with a single newline
                        text = text.replace(/\n{2,}/g, '\n');
                        // Trim leading/trailing whitespace
                        text = text.trim();

                        // Insert plain text at cursor position
                        document.execCommand('insertText', false, text);
                    }
                });

                var $footer = $('<div class="modal-footer"></div>');
                var $statusMsg = $('<div class="status-message"></div>');
                var $cancelBtn = $('<button class="modal-btn cancel-btn">Cancel</button>');
                $cancelBtn.on('click', function() {
                    closeModalAndCleanup();
                });
                var $saveBtn = $('<button class="modal-btn save-btn">Save</button>');
                $saveBtn.on('click', function() {
                    // Get edited content
                    var editedContent = $editor.html();

                    // === CLEANUP: Remove trailing empty divs with just <br> ===
                    var $tempDiv = $('<div>').html(editedContent);

                    // Remove trailing empty divs
                    var $children = $tempDiv.children();
                    for (var i = $children.length - 1; i >= 0; i--) {
                        var $child = $children.eq(i);
                        var html = $child.html();
                        if ($child.is('div') && (html === '<br>' || html === '<br/>' || html.trim() === '')) {
                            $child.remove();
                        } else {
                            break; // Stop at first non-empty element
                        }
                    }

                    editedContent = $tempDiv.html();

                    // === CONVERT SECTION CONTENT BADGE TO PLACEHOLDER ===
                    $tempDiv = $('<div>').html(editedContent);

                    // Find the qlik-section-content badge (if still exists - user might have deleted it)
                    var $sectionBadge = $tempDiv.find('.qlik-section-content');

                    if ($sectionBadge.length > 0) {
                        // Badge exists - replace with placeholder marker
                        // This will be re-rendered on load with fresh data
                        $sectionBadge.replaceWith('⟨⟨SECTION_CONTENT⟩⟩');
                    }

                    editedContent = $tempDiv.html();

                    $saveBtn.prop('disabled', true);
                    $statusMsg.text('Checking for conflicts...').removeClass('error success').css('color', '#666');

                    // Set flag to prevent paint() during save (prevents Error code 16 crash)
                    savingInProgress = true;

                    // Get appId from layout - no need for Qlik API calls
                    var app = qlik.currApp();
                    var appId = app.id;

                    // === CONFLICT DETECTION: Check if PK file was modified by someone else ===
                    loadModifiedSections(appId, pkValue, function(currentData) {
                        var currentTimestamp = null;

                        // Get the current timestamp from the PK file (NEW STRUCTURE: direct sections)
                        if (currentData &&
                            currentData.sections &&
                            currentData.sections[section.label]) {
                            currentTimestamp = currentData.sections[section.label].lastModified;
                            var currentModifiedBy = currentData.sections[section.label].modifiedBy;

                            // Check if timestamp has changed since modal was opened
                            if (modalBaselineTimestamp && currentTimestamp !== modalBaselineTimestamp) {
                                // CONFLICT DETECTED!
                        var conflictMsg = '⚠️ CONFLICT: This section was modified by ' + currentModifiedBy +
                                    ' at ' + new Date(currentTimestamp).toLocaleString() +
                                    ' while you were editing.\n\n' +
                                    'Your changes cannot be saved to prevent data loss.\n\n' +
                                    'Please:\n' +
                                    '1. Copy your changes to clipboard\n' +
                                    '2. Close this modal\n' +
                                    '3. Reopen to see their changes\n' +
                                    '4. Merge your changes manually';

                                $statusMsg.html(conflictMsg.replace(/\n/g, '<br>')).addClass('error');
                                $saveBtn.prop('disabled', false);
                        savingInProgress = false;
                        
                        return; // ABORT SAVE
                            }
                        }

                        // No conflict - proceed with save
                        
                        $statusMsg.text('Saving...').removeClass('error success').css('color', '#666');

                        // Save directly to SharePoint (no spaceId or connectionId needed anymore)
                        saveModifiedSections(appId, null, null, pkValue, section.label, editedContent, function(success, errorMsg) {

                        if (success) {
                            $statusMsg.text('✅ Saved successfully! Updating...').addClass('success');

                            // Reload fresh data from SharePoint to ensure cache is in sync
                            // This is critical to preserve all sections when multiple edits are made
                            var currentPkValue = pkValue; // Capture for closure

                            // Temporarily mark cache as not loaded to force reload from SharePoint
                            modifiedSectionsCache.loaded = false;

                            loadModifiedSections(appId, currentPkValue, function(freshData) {
                                // Update cache with complete fresh data from SharePoint
                                if (freshData && freshData.sections) {
                                    modifiedSectionsCache.data = freshData;
                                    modifiedSectionsCache.loaded = true;
                                    modifiedSectionsCache.pkValue = currentPkValue;
                                } else {
                                    // If load fails, do a partial cache update to at least reflect this section
                                    if (!modifiedSectionsCache.data) {
                                        modifiedSectionsCache.data = {
                                            sections: {},
                                            appId: appId,
                                            pkValue: currentPkValue,
                                            created: new Date().toISOString()
                                        };
                                    }
                                    if (!modifiedSectionsCache.data.sections) {
                                        modifiedSectionsCache.data.sections = {};
                                    }
                                    var modifiedByDisplay = currentUserName !== 'Unknown User' && currentUserEmail !== 'user@qlik.com'
                                        ? currentUserName + ' (' + currentUserEmail + ')'
                                        : currentUserEmail;

                                    modifiedSectionsCache.data.sections[section.label] = {
                                        content: editedContent,
                                        lastModified: new Date().toISOString(),
                                        modifiedBy: modifiedByDisplay
                                    };
                                    modifiedSectionsCache.data.lastModified = new Date().toISOString();
                                    modifiedSectionsCache.loaded = true;
                                    modifiedSectionsCache.pkValue = currentPkValue;
                                }

                                // Re-render the section with new content
                                var $contentElement = sectionData.$content;
                            if ($contentElement && $contentElement.length > 0) {
                                // Process the placeholder if it exists
                                var finalContent = editedContent;

                                // Strip any existing label+separator from edited content before re-rendering
                                if (section.showLabel && section.label) {
                                    var baseFontSize = parseInt(layout.fontSize || '14', 10);
                                    var labelSizeOffset = parseInt(layout.labelSizeOffset || '4', 10);
                                    var labelSize = (baseFontSize + labelSizeOffset) + 'px';
                                    var labelColor = section.labelColor || '#1a1a1a';

                                    // Remove label div if present
                                    var labelPatterns = [
                                        '<div style="font-size: ' + labelSize + '; color: ' + labelColor,
                                        '<div style="font-size:' + labelSize + '; color:' + labelColor
                                    ];

                                    for (var pi = 0; pi < labelPatterns.length; pi++) {
                                        if (finalContent.indexOf(labelPatterns[pi]) === 0) {
                                            var labelEnd = finalContent.indexOf(section.label + '</div>');
                                            if (labelEnd > -1) {
                                                finalContent = finalContent.substring(labelEnd + (section.label + '</div>').length);

                                                // Also remove separator if present
                                                if (finalContent.indexOf('<div style="height: 2px;') === 0 || finalContent.indexOf('<div style="height: 1px;') === 0) {
                                                    var sepEnd = finalContent.indexOf('</div>');
                                                    if (sepEnd > -1) {
                                                        finalContent = finalContent.substring(sepEnd + 6);
                                                    }
                                                }
                                                break;
                                            }
                                        }
                                    }
                                }

                                // === CHECK FOR SECTION CONTENT PLACEHOLDER ===
                                if (finalContent.indexOf('⟨⟨SECTION_CONTENT⟩⟩') > -1) {
                                    // Re-render the original template with fresh data
                                    var freshTemplateContent = section.markdownText || '';

                                    // Get the data from sectionData parameter
                                    var data = sectionData ? sectionData.data : null;
                                    var numDimensions = sectionData ? sectionData.numDimensions : 0;
                                    var numMeasures = sectionData ? sectionData.numMeasures : 0;
                                    var itemMapping = sectionData ? sectionData.itemMapping : {};

                                    // STEP 1: Convert {{[Master Item Name]}} to {{dim1}}, {{measure1}} format
                                    Object.keys(itemMapping).forEach(function(itemName) {
                                        var mapping = itemMapping[itemName];
                                        var masterPattern = '{{[' + itemName + ']}}';
                                        var indexPattern;

                                        if (mapping.type === 'dim') {
                                            indexPattern = '{{dim' + (mapping.index + 1) + '}}';
                                        } else {
                                            indexPattern = '{{measure' + (mapping.index + 1) + '}}';
                                        }

                                        var regex = new RegExp(masterPattern.replace(/[{}[\]]/g, '\\$&'), 'g');
                                        freshTemplateContent = freshTemplateContent.replace(regex, indexPattern);
                                    });

                                    // STEP 2: Replace placeholders with actual values (outside iterating tags)
                                    if (data && data.qMatrix && data.qMatrix.length > 0) {
                                        function isInsideIteratingTagModal(text, matchIndex) {
                                            var iteratingTags = ['list', 'table', 'concat', 'grid', 'kpi'];
                                            var before = text.substring(0, matchIndex);
                                            var lastOpenMatch = null;
                                            var lastOpenIndex = -1;
                                            iteratingTags.forEach(function(tag) {
                                                var regex = new RegExp('#\\[' + tag + '[^\\]]*\\]', 'g');
                                                var match;
                                                while ((match = regex.exec(before)) !== null) {
                                                    if (match.index > lastOpenIndex) {
                                                        lastOpenIndex = match.index;
                                                        lastOpenMatch = tag;
                                                    }
                                                }
                                            });
                                            if (lastOpenMatch) {
                                                var afterOpen = before.substring(lastOpenIndex);
                                                var closeRegex = new RegExp('#\\[\\/' + lastOpenMatch + '\\]');
                                                if (!closeRegex.test(afterOpen)) {
                                                    return true;
                                                }
                                            }
                                            return false;
                                        }

                                        for (var dimIdx = 0; dimIdx < numDimensions; dimIdx++) {
                                            var placeholder = '{{dim' + (dimIdx + 1) + '}}';
                                            var regex = new RegExp(placeholder.replace(/[{}]/g, '\\$&'), 'g');
                                            var cell = data.qMatrix[0][dimIdx];
                                            var value = cell ? cell.qText : 'N/A';
                                            freshTemplateContent = freshTemplateContent.replace(regex, function(match, offset) {
                                                return isInsideIteratingTagModal(freshTemplateContent, offset) ? match : value;
                                            });
                                        }

                                        for (var meaIdx = 0; meaIdx < numMeasures; meaIdx++) {
                                            var placeholder = '{{measure' + (meaIdx + 1) + '}}';
                                            var regex = new RegExp(placeholder.replace(/[{}]/g, '\\$&'), 'g');
                                            var cell = data.qMatrix[0][numDimensions + meaIdx];
                                            var value = cell ? cell.qText : 'N/A';
                                            freshTemplateContent = freshTemplateContent.replace(regex, function(match, offset) {
                                                return isInsideIteratingTagModal(freshTemplateContent, offset) ? match : value;
                                            });
                                        }
                                    } else {
                                        // No data - replace with [No data] message
                                        freshTemplateContent = freshTemplateContent.replace(/\{\{dim\d+\}\}/g, '<span style="color: #999; font-style: italic;">[No data]</span>');
                                        freshTemplateContent = freshTemplateContent.replace(/\{\{measure\d+\}\}/g, '<span style="color: #999; font-style: italic;">[No data]</span>');
                                        freshTemplateContent = freshTemplateContent.replace(/\{\{\[([^\]]+)\]\}\}/g, '<span style="color: #f57c00; font-style: italic;" title="Master item not found">[Master item &quot;$1&quot; not found]</span>');
                                    }

                                    // STEP 3: Process content tags in the template
                                    if (freshTemplateContent.indexOf('#[') !== -1) {
                                        freshTemplateContent = processContentTags(freshTemplateContent, data, numDimensions, numMeasures, layout, null, itemMapping);
                                    }

                                    // STEP 4: Parse the template markdown
                                    var freshRendered = parseMarkdown(freshTemplateContent);

                                    // STEP 5: Replace the placeholder with fresh content
                                    finalContent = finalContent.replace(/⟨⟨SECTION_CONTENT⟩⟩/g, freshRendered);
                                }

                                // editedContent is already HTML from the contenteditable editor
                                // Add label back if showLabel is enabled
                        var contentWithLabel = finalContent;
                        if (section.showLabel && section.label) {
                                    var labelHtml = '';
                                    var labelStyle = section.labelStyle || 'bold';
                                    // Calculate label size: default font + offset
                                    var baseFontSize = parseInt(layout.fontSize || '14', 10);
                                    var labelSizeOffset = parseInt(layout.labelSizeOffset || '4', 10);
                                    var labelSize = (baseFontSize + labelSizeOffset) + 'px';
                                    var labelColor = section.labelColor || '#1a1a1a';
                                    var labelSeparator = (section.labelSeparator === true || section.labelSeparator === 'true');

                                    var styleAttr = 'font-size: ' + labelSize + '; color: ' + labelColor + '; margin: 0; line-height: 1.3; padding: 0 0 1px 0;';

                                    if (labelStyle === 'bold') {
                                        styleAttr += ' font-weight: 600;';
                                    } else if (labelStyle === 'italic') {
                                        styleAttr += ' font-style: italic;';
                                    }

                                    labelHtml = '<div style="' + styleAttr + '">' + section.label + '</div>';

                                    if (labelSeparator) {
                                        labelHtml += '<div style="height: 2px; background: linear-gradient(to right, ' + labelColor + ', transparent); margin-bottom: 5px; opacity: 0.5; position: relative; z-index: 1;"></div>';
                                    }

                                    contentWithLabel = labelHtml + finalContent;
                                }
                                $contentElement.html(contentWithLabel);
                                } else {

                                }

                                // Close modal
                                closeModalAndCleanup();

                                // Reset saving flag after brief delay to ensure callbacks have settled
                                setTimeout(function() {
                                    savingInProgress = false;
                                }, 1000);
                            }); // End loadModifiedSections callback (cache refresh after save)
                        } else {
                            $statusMsg.text('❌ Save failed: ' + errorMsg).addClass('error');
                            $saveBtn.prop('disabled', false);
                            savingInProgress = false;
                        }
                    }); // End saveModifiedSections callback
                    }); // End loadModifiedSections callback (conflict detection)
                }); // End $saveBtn.on('click')

                $footer.append($statusMsg);
                $footer.append($cancelBtn);
                $footer.append($saveBtn);

                $modal.append($header);
                $modal.append($body);
                $modal.append($footer);
                $overlay.append($modal);

                // Store the token and generic object for this modal
                currentModalPKToken = pkToken;
                currentModalPKObject = pkCheckResult.obj;

                $('body').append($overlay);

                

                // Close on overlay click
                $overlay.on('click', function(e) {
                    if (e.target === $overlay[0]) {
                        closeModalAndCleanup();
                    }
                });
                } // End openModalWithFreshContent function
            }); // End loadModifiedSections callback (fresh data load)
        }); // End getCurrentPKValue callback
    }

    // Helper function to close modal and cleanup PK object
    function closeModalAndCleanup() {
        $('.edit-modal-overlay').remove();

        // Cleanup dynamic CSS added for editor lists
        $('style[id^="list-style-editor-"]').remove();

        // === FAST POLLING: Disable when modal closes ===
        if (modalIsOpen) {
            modalIsOpen = false;
            modalBaselineTimestamp = null;
            modalSectionLabel = null;
            modalPKValue = null;

        }

        // Close and cleanup the PK generic object for this modal
        if (currentModalPKObject) {
            try {
                if (typeof currentModalPKObject.close === 'function') {
                    // Direct close method exists
                    currentModalPKObject.close();
                } else if (currentModalPKObject.then && typeof currentModalPKObject.then === 'function') {
                    // It's a promise - resolve it first then close
                    currentModalPKObject.then(function(obj) {
                        if (obj && typeof obj.close === 'function') {
                            obj.close();
                        }
                    });
                }
            } catch (e) {
                
            }
            currentModalPKObject = null;
        }
        currentModalPKToken = null;
    }

    // Throttled 
    function throttledWarn(message) {
        var now = Date.now();
        if (message === lastWarningMessage && (now - lastWarningTime) < 2000) {
            return; // Suppress repeated warning within 2 seconds
        }
        lastWarningMessage = message;
        lastWarningTime = now;
        
    }

    function getCurrentPKValue(pkField, callback) {
        var app = qlik.currApp();

        // Generate unique token for this PK check
        var myToken = ++currentPKCheckToken;

        // Create a single generic object with all needed expressions
        // Use Only() to let Qlik handle unique value extraction efficiently
        var checkObjPromise = app.createGenericObject({
            allSelectedFields: {
                qStringExpression: "GetCurrentSelections('||', '=')"
            },
            pkSelection: {
                qStringExpression: "GetFieldSelections('" + pkField + "')"
            },
            pkCount: {
                qStringExpression: "GetSelectedCount('" + pkField + "')"
            },
            pkOnlyValue: {
                qStringExpression: "Only([" + pkField + "])"
            }
        }, function(reply) {
            // CRITICAL: Check if this callback is stale
            if (myToken !== currentPKCheckToken) {
                return;
            }

            var allSelections = reply.allSelectedFields;
            var count = parseInt(reply.pkCount) || 0;
            var value = reply.pkSelection;
            var onlyValue = reply.pkOnlyValue;

            // Only() returns the value if there's exactly one unique value, otherwise null
            // This handles the case where hypercube has 3 duplicate rows but 1 unique PK
            if (onlyValue && onlyValue !== '-') {
                callback(onlyValue, null, myToken);
                return;
            }

            // Only() returned null or '-', meaning either:
            // - Multiple unique values exist
            // - No values available
            // - Field doesn't exist

            // Check if user has made an explicit selection
            if (count > 1) {
                
                callback(null, 'Multiple values selected in ' + pkField + '. Please select exactly one.', myToken);
                return;
            }

            if (count === 1 && value && value !== '-') {
                callback(value, null, myToken);
                return;
            }

            // Fallback: Try to parse from GetCurrentSelections (handles implicit selections)
            if (allSelections && allSelections.indexOf(pkField + '=') !== -1) {
                var parts = allSelections.split('||');
                for (var i = 0; i < parts.length; i++) {
                    var part = parts[i];
                    if (part.indexOf(pkField + '=') === 0) {
                        var extractedValue = part.substring(pkField.length + 1); // +1 for '='
                        callback(extractedValue, null, myToken);
                        return;
                    }
                }
            }

            // No selection made - inform user they need to select a value
            
            callback(null, 'No value selected in ' + pkField + '. Please select exactly one value.', myToken);
        });

        // The promise will resolve to the actual object with the close method
        // We store the promise and close it via promise resolution

        // Store the promise so we can close it if needed
        pkCheckObjects.push(checkObjPromise);

        // Keep only last 5 objects, close older ones
        if (pkCheckObjects.length > 5) {
            var oldObj = pkCheckObjects.shift();
            try {
                if (oldObj && typeof oldObj.close === 'function') {
                    oldObj.close();
                }
            } catch (e) {
                // Ignore errors
            }
        }

        // Return the promise - we'll close it via .then() to get the actual object
        return { obj: checkObjPromise, token: myToken };
    }

    function renderSectionsWithData($element, layout, sectionsData, self, instanceId) {
        // Capture current render token to detect stale async callbacks
        var renderToken = currentRenderToken[instanceId];

        // Close any open modals when rendering starts (selection changed)
        closeModalAndCleanup();

        // Ensure currentModel is set
        if (!currentModel) {
            currentModel = self.backendApi.model;
        }

        // === MULTI-USER: Only load saved content if feature is enabled ===
        if (MULTI_USER_ENABLED === 1 && layout.pkField) {
            // Get PK value fresh each time (no caching)
            getCurrentPKValue(layout.pkField, function(pkValue, error) {
                // Check if render is still current
                if (currentRenderToken[instanceId] !== renderToken) {
                    return;
                }

                // If valid PK, load modified sections
                if (!error && pkValue && pkValue !== '-') {
                    modifiedSectionsCache.pkValue = pkValue;

                    var app = qlik.currApp();
                    var appId = app.id;

                    // Update loading message when loading from SharePoint
                    updateLoadingMessage($element, 'Fetching saved content...', 'Loading from SharePoint for PK: ' + pkValue);

                    loadModifiedSections(appId, pkValue, function(data) {
                        // Check if this render is still current
                        if (currentRenderToken[instanceId] !== renderToken) {
                            return;
                        }
                        modifiedSectionsCache.data = data;
                        modifiedSectionsCache.loaded = true;
                        modifiedSectionsCache.pkValue = pkValue;

                        // Update loading message before final render
                        updateLoadingMessage($element, 'Building sections...', 'Rendering content');

                        renderSectionsUI($element, layout, sectionsData, self, instanceId);
                    });
                } else {
                    // No valid single PK value - proceed with normal rendering
                    modifiedSectionsCache.loaded = false;

                    // Update loading message before final render
                    updateLoadingMessage($element, 'Building sections...', 'Rendering content');

                    renderSectionsUI($element, layout, sectionsData, self, instanceId);
                }
            });
        } else {
            // Multi-user disabled or no PK field - proceed with normal rendering
            modifiedSectionsCache.loaded = false;

            // Update loading message before final render
            updateLoadingMessage($element, 'Building sections...', 'Rendering content');

            renderSectionsUI($element, layout, sectionsData, self, instanceId);
        }
    }

    // GUI Content Editor Modal with Master Items Support - Phase 2
    // NOTE: The onSave callback MUST use vis.model.setProperties() to persist changes
    //       See line ~530 for the proper implementation (Fixed: 2026-03-19)
    function showGUIEditorModal(section, sectionData, layout, onSave) {
        

        var app = qlik.currApp();
        var masterItems = []; // Will store {name, type: 'dimension'|'measure'}

        // Create modal overlay
        var $overlay = $('<div class="edit-modal-overlay"></div>');
        var $modal = $('<div class="edit-modal" style="max-width: 1000px; height: 85vh;"></div>');

        // Add custom CSS for better cursor visibility with block cursor
        var $customCSS = $('<style>' +
            '.gui-editor:focus { outline: 2px solid #1976d2; outline-offset: -2px; }' +
            '.gui-editor { caret-color: #1976d2; caret-shape: block; }' + // Block cursor in Firefox
            '.gui-editor .tag-badge { pointer-events: auto; }' +
            '.gui-editor .master-item-badge { margin-right: 10px !important; }' + // Extra space after badges
            '</style>');
        $modal.append($customCSS);

        // Modal header
        var $header = $('<div class="modal-header"></div>');
        $header.append('<h3>📝 Visual Content Editor</h3>');
        $header.append('<div style="font-size: 12px; color: #666; margin-top: 4px;">Edit ' + (section.label || 'Section') + '</div>');

        var $closeBtn = $('<button class="modal-close-btn">&times;</button>');
        $closeBtn.on('click', function() {
            $overlay.remove();
        });
        $header.append($closeBtn);

        // Modal body
        var $body = $('<div class="modal-body" style="display: flex; flex-direction: column; padding: 20px; height: calc(100% - 140px); position: relative;"></div>');

        // Info message
        var $info = $('<div style="background: #e3f2fd; border-left: 4px solid #1976d2; padding: 12px; margin-bottom: 15px; border-radius: 4px;">' +
            '<strong>💡 Tips:</strong><br>' +
            '• Click <strong>📦 Styled Box</strong> to insert formatted container (WYSIWYG mode)<br>' +
            '• Use <strong>⬅ ↔ ➡</strong> buttons to align text inside styled boxes<br>' +
            '• Use toolbar buttons (Bold, Color, Size) to format text, then save<br>' +
            '• Press <strong>Enter</strong> for line breaks<br>' +
            '• Saves as: <code>{bg:#color padding:10px align:left}content{/bg}</code><br>' +
            '<strong>Note:</strong> Manual syntax for inline colors: <code>{#fff:size:18px:bold:text}</code> (color comes FIRST, no "color:" prefix!)' +
            '</div>');

        // Component Toolbar
        var $componentToolbar = $('<div style="background: #fafafa; border: 1px solid #d0d0d0; border-radius: 4px; padding: 8px; margin-bottom: 10px; display: flex; gap: 5px; flex-wrap: wrap; align-items: center;"></div>');

        // Insert Master Item button (prominent, first button)
        var $insertMasterItemBtn = $('<button style="' +
            'background: #1976d2; ' +
            'color: white; ' +
            'border: none; ' +
            'padding: 8px 12px; ' +
            'border-radius: 4px; ' +
            'cursor: pointer; ' +
            'font-size: 13px; ' +
            'font-weight: 600; ' +
            'display: flex; ' +
            'align-items: center; ' +
            'gap: 6px; ' +
            'margin-right: 15px;' +
        '">📊 Insert Master Item</button>');

        $insertMasterItemBtn.on('click', function() {
            $editor.focus();
            showDropdown('');
        });

        $insertMasterItemBtn.hover(
            function() { $(this).css('background', '#1565c0'); },
            function() { $(this).css('background', '#1976d2'); }
        );

        $componentToolbar.append($insertMasterItemBtn);

        var $toolbarLabel = $('<span style="font-size: 11px; color: #666; font-weight: 600; margin-right: 10px;">COMPONENTS:</span>');
        $componentToolbar.append($toolbarLabel);

        // Component buttons (only Styled Box for now - WYSIWYG mode)
        var components = [
            { icon: '📦', name: 'styledbox', label: 'Styled Box' }
        ];

        components.forEach(function(comp) {
            var $btn = $('<button class="component-btn" title="Insert ' + comp.label + '" style="' +
                'background: white; ' +
                'border: 1px solid #d0d0d0; ' +
                'border-radius: 3px; ' +
                'padding: 5px 10px; ' +
                'cursor: pointer; ' +
                'font-size: 16px; ' +
                'transition: all 0.2s;' +
                '">' + comp.icon + '</button>');

            $btn.hover(
                function() { $(this).css({ 'background': '#e3f2fd', 'border-color': '#1976d2' }); },
                function() { $(this).css({ 'background': 'white', 'border-color': '#d0d0d0' }); }
            );

            $btn.on('click', function() {
                insertComponentTag(comp.name);
            });

            $componentToolbar.append($btn);
        });

        // Formatting Toolbar - Match Edit Data Modal exactly (same order: B, I, U, Size, Color)
        var $formatToolbar = $('<div class="editor-toolbar" style="background: #f5f5f5; border: 1px solid #d0d0d0; border-radius: 4px; padding: 8px; margin-bottom: 10px; display: flex; gap: 4px; flex-wrap: wrap; align-items: center;"></div>');

        // Helper function to create toolbar button (matches Edit modal)
        function createToolbarBtn(label, command, value) {
            var $btn = $('<button type="button" style="' +
                'padding: 6px 12px; ' +
                'border: 1px solid #ccc; ' +
                'background: white; ' +
                'border-radius: 3px; ' +
                'cursor: pointer; ' +
                'font-size: 14px; ' +
                'min-width: 36px; ' +
                'height: 34px;' +
                '"></button>');
            $btn.html(label);

            $btn.hover(
                function() { $(this).css({ 'background': '#e0e0e0', 'border-color': '#999' }); },
                function() { $(this).css({ 'background': 'white', 'border-color': '#ccc' }); }
            );

            $btn.on('mousedown', function(e) {
                e.preventDefault(); // Prevent losing focus
            });

            $btn.on('click', function(e) {
                e.preventDefault();
                if (command === 'custom') {
                    value(); // Call custom function
                } else {
                    document.execCommand(command, false, value);
                }
                $editor.focus();
            });

            return $btn;
        }

        // Text formatting buttons (B, I, U)
        $formatToolbar.append(createToolbarBtn('<b>B</b>', 'bold'));
        $formatToolbar.append(createToolbarBtn('<i>I</i>', 'italic'));
        $formatToolbar.append(createToolbarBtn('<u>U</u>', 'underline'));

        // Separator
        $formatToolbar.append('<div style="width: 1px; height: 24px; background: #ccc; margin: 0 4px;"></div>');

        // Alignment buttons - custom handlers for styled boxes
        var $alignLeftBtn = $('<button type="button" title="Align Left" style="padding: 6px 12px; border: 1px solid #ccc; background: white; border-radius: 3px; cursor: pointer; font-size: 14px; min-width: 36px; height: 34px;">⬅</button>');
        $alignLeftBtn.hover(
            function() { $(this).css({ 'background': '#e0e0e0', 'border-color': '#999' }); },
            function() { $(this).css({ 'background': 'white', 'border-color': '#ccc' }); }
        );
        $alignLeftBtn.on('click', function(e) {
            e.preventDefault();
            var selection = window.getSelection();
            if (selection.rangeCount > 0) {
                var $box = $(selection.anchorNode).closest('.wysiwyg-component[data-component-type="styledbox"]');
                if ($box.length > 0) {
                    var currentStyle = $box.attr('style') || '';
                    currentStyle = currentStyle.replace(/text-align:[^;]+;?/gi, '');
                    $box.attr('style', currentStyle + ' text-align: left !important;');
                    var attrs = $box.attr('data-attrs') || '';
                    attrs = attrs.replace(/align:[^\s]+\s*/g, '');
                    attrs += ' align:left';
                    $box.attr('data-attrs', attrs.trim());
                } else {
                    document.execCommand('justifyLeft', false, null);
                }
            }
            $editor.focus();
        });

        var $alignCenterBtn = $('<button type="button" title="Align Center" style="padding: 6px 12px; border: 1px solid #ccc; background: white; border-radius: 3px; cursor: pointer; font-size: 14px; min-width: 36px; height: 34px;">↔</button>');
        $alignCenterBtn.hover(
            function() { $(this).css({ 'background': '#e0e0e0', 'border-color': '#999' }); },
            function() { $(this).css({ 'background': 'white', 'border-color': '#ccc' }); }
        );
        $alignCenterBtn.on('click', function(e) {
            e.preventDefault();
            var selection = window.getSelection();
            if (selection.rangeCount > 0) {
                var $box = $(selection.anchorNode).closest('.wysiwyg-component[data-component-type="styledbox"]');
                if ($box.length > 0) {
                    var currentStyle = $box.attr('style') || '';
                    currentStyle = currentStyle.replace(/text-align:[^;]+;?/gi, '');
                    $box.attr('style', currentStyle + ' text-align: center !important;');
                    var attrs = $box.attr('data-attrs') || '';
                    attrs = attrs.replace(/align:[^\s]+\s*/g, '');
                    attrs += ' align:center';
                    $box.attr('data-attrs', attrs.trim());
                } else {
                    document.execCommand('justifyCenter', false, null);
                }
            }
            $editor.focus();
        });

        var $alignRightBtn = $('<button type="button" title="Align Right" style="padding: 6px 12px; border: 1px solid #ccc; background: white; border-radius: 3px; cursor: pointer; font-size: 14px; min-width: 36px; height: 34px;">➡</button>');
        $alignRightBtn.hover(
            function() { $(this).css({ 'background': '#e0e0e0', 'border-color': '#999' }); },
            function() { $(this).css({ 'background': 'white', 'border-color': '#ccc' }); }
        );
        $alignRightBtn.on('click', function(e) {
            e.preventDefault();
            var selection = window.getSelection();
            if (selection.rangeCount > 0) {
                var $box = $(selection.anchorNode).closest('.wysiwyg-component[data-component-type="styledbox"]');
                if ($box.length > 0) {
                    var currentStyle = $box.attr('style') || '';
                    currentStyle = currentStyle.replace(/text-align:[^;]+;?/gi, '');
                    $box.attr('style', currentStyle + ' text-align: right !important;');
                    var attrs = $box.attr('data-attrs') || '';
                    attrs = attrs.replace(/align:[^\s]+\s*/g, '');
                    attrs += ' align:right';
                    $box.attr('data-attrs', attrs.trim());
                } else {
                    document.execCommand('justifyRight', false, null);
                }
            }
            $editor.focus();
        });

        $formatToolbar.append($alignLeftBtn);
        $formatToolbar.append($alignCenterBtn);
        $formatToolbar.append($alignRightBtn);

        // Separator
        $formatToolbar.append('<div style="width: 1px; height: 24px; background: #ccc; margin: 0 4px;"></div>');

        // Font size dropdown (matches Edit modal exactly)
        var $fontSizeSelect = $('<select title="Font Size" style="padding: 6px 8px; border: 1px solid #ccc; border-radius: 3px; background: white; cursor: pointer; font-size: 14px; height: 34px;"></select>');
        $fontSizeSelect.append('<option value="">Font Size</option>');
        var sizes = ['10', '12', '14', '16', '18', '20', '22', '24', '28', '32', '36', '40', '48', '56', '64', '72'];
        sizes.forEach(function(size) {
            $fontSizeSelect.append('<option value="' + size + 'px">' + size + 'px</option>');
        });

        $fontSizeSelect.on('change', function() {
            var size = $(this).val();
            if (size) {
                var selection = window.getSelection();
                if (selection.rangeCount > 0 && !selection.isCollapsed) {
                    var range = selection.getRangeAt(0);
                    var span = document.createElement('span');
                    span.style.fontSize = size;
                    span.setAttribute('data-format-type', 'size');
                    span.setAttribute('data-format-value', size);
                    try {
                        range.surroundContents(span);
                    } catch(e) {
                        // Fallback if surroundContents fails
                        var contents = range.extractContents();
                        span.appendChild(contents);
                        range.insertNode(span);
                    }
                }
                $(this).val('');
                $editor.focus();
            }
        });

        $fontSizeSelect.hover(
            function() { $(this).css('border-color', '#999'); },
            function() { $(this).css('border-color', '#ccc'); }
        );

        $formatToolbar.append($fontSizeSelect);

        // Separator
        $formatToolbar.append('<div style="width: 1px; height: 24px; background: #ccc; margin: 0 4px;"></div>');

        // Color picker (matches Edit modal)
        var savedColorSelection = null;
        var $colorPickerWrapper = $('<div style="display: inline-flex; align-items: center; gap: 4px;"></div>');
        var $colorPicker = $('<input type="color" value="#333333" title="Text Color" style="width: 40px; height: 34px; border: 1px solid #ccc; border-radius: 3px; cursor: pointer; background: white; padding: 2px;">');

        $colorPicker.on('mousedown', function(e) {
            var selection = window.getSelection();
            if (selection.rangeCount > 0) {
                savedColorSelection = selection.getRangeAt(0);
            }
        });

        $colorPicker.on('input change', function() {
            var color = $(this).val();

            if (savedColorSelection) {
                $editor.focus();
                var selection = window.getSelection();
                selection.removeAllRanges();
                selection.addRange(savedColorSelection);
            }

            document.execCommand('foreColor', false, color);
            $editor.focus();
        });

        $colorPicker.hover(
            function() { $(this).css('border-color', '#999'); },
            function() { $(this).css('border-color', '#ccc'); }
        );

        $colorPickerWrapper.append($colorPicker);
        $colorPickerWrapper.append('<span style="font-size: 11px; color: #666;">Color</span>');
        $formatToolbar.append($colorPickerWrapper);

        // Separator
        $formatToolbar.append('<div style="width: 1px; height: 24px; background: #ccc; margin: 0 4px;"></div>');

        // Box background color picker (works when editing inside a styled box component)
        var $boxBgColorWrapper = $('<div style="display: inline-flex; align-items: center; gap: 4px;"></div>');
        var $boxBgColorPicker = $('<input type="color" value="#006580" title="Box Background Color (works inside styled box)" style="width: 40px; height: 34px; border: 1px solid #ccc; border-radius: 3px; cursor: pointer; background: white; padding: 2px;">');

        $boxBgColorPicker.on('input change', function() {
            var color = $(this).val();

            // Find the styled box component that contains the current selection
            var selection = window.getSelection();
            if (selection.rangeCount > 0) {
                var node = selection.anchorNode;
                var $boxComponent = $(node).closest('.wysiwyg-component[data-component-type="styledbox"]');

                if ($boxComponent.length > 0) {
                    // Update the box's background color
                    $boxComponent.css('background-color', color);

                    // Update the attrs to store the new bg color
                    var currentAttrs = $boxComponent.attr('data-attrs') || '';
                    // Remove existing bg color (starts with # after "bg:")
                    currentAttrs = currentAttrs.replace(/#[0-9a-fA-F]{6}\s*/gi, color + ' ');
                    if (currentAttrs.indexOf(color) === -1) {
                        // If color wasn't in attrs yet, add it at the start
                        currentAttrs = color + ' ' + currentAttrs;
                    }
                    $boxComponent.attr('data-attrs', currentAttrs.trim());

                    $editor.focus();
                } else {
                    alert('Please place your cursor inside a styled box to change its background color.');
                }
            }
        });

        $boxBgColorPicker.hover(
            function() { $(this).css('border-color', '#999'); },
            function() { $(this).css('border-color', '#ccc'); }
        );

        $boxBgColorWrapper.append($boxBgColorPicker);
        $boxBgColorWrapper.append('<span style="font-size: 11px; color: #666;">Box BG</span>');
        $formatToolbar.append($boxBgColorWrapper);

        // Editor container with contenteditable
        var $editorContainer = $('<div style="flex: 1; border: 1px solid #d0d0d0; border-radius: 4px; overflow: hidden; display: flex; flex-direction: column; position: relative;"></div>');

        var $editorToolbar = $('<div style="background: #f5f5f5; padding: 8px; border-bottom: 1px solid #d0d0d0; font-size: 12px; color: #666;">' +
            '<span style="color: #009845;">Dimensions: <strong>' + masterItems.filter(function(m) { return m.type === 'dimension'; }).length + '</strong></span> ' +
            '<span style="margin-left: 15px; color: #1976d2;">Measures: <strong>' + masterItems.filter(function(m) { return m.type === 'measure'; }).length + '</strong></span>' +
            '</div>');

        // Apply default font, size, and color from appearance settings
        var defaultFont = layout.fontFamily || 'Arial, sans-serif';
        var defaultSize = layout.fontSize || '14';
        var defaultColor = '#333333'; // Dark grey to match component default
        var $editor = $('<div class="gui-editor" contenteditable="true" style="flex: 1; padding: 15px; overflow-y: auto; font-family: ' + defaultFont + '; font-size: ' + defaultSize + 'px; color: ' + defaultColor + '; line-height: 1.8; background: white; caret-color: #1976d2; caret-shape: block;"></div>');

        // Handle Enter key to insert <br> instead of <div> - use event delegation for dynamic components
        $editor.on('keydown', function(e) {
            if (e.keyCode === 13 && !e.shiftKey) { // Enter key (Shift+Enter for default behavior)
                e.preventDefault();
                e.stopPropagation();

                var selection = window.getSelection();
                if (selection.rangeCount > 0) {
                    var range = selection.getRangeAt(0);

                    // Insert <br> tag
                    var br = document.createElement('br');
                    range.deleteContents();
                    range.insertNode(br);

                    // Move cursor after the <br>
                    range.setStartAfter(br);
                    range.collapse(true);

                    // Check if we're at the end of a container - if so, add extra <br> for spacing
                    // This prevents the need to press Enter twice
                    var container = range.startContainer;
                    if (container.nodeType === Node.TEXT_NODE) {
                        container = container.parentNode;
                    }

                    // If cursor is at the end, insert second <br> to ensure visible line
                    var atEnd = !range.startContainer.nextSibling ||
                               (range.startContainer.nodeType === Node.TEXT_NODE &&
                                range.startOffset >= range.startContainer.length);

                    if (atEnd) {
                        var br2 = document.createElement('br');
                        range.insertNode(br2);
                        range.setStartAfter(br2);
                        range.collapse(true);
                    }

                    selection.removeAllRanges();
                    selection.addRange(range);
                }
                return false;
            }
        });

        // Handle Enter key inside wysiwyg components (delegated)
        $editor.on('keydown', '.wysiwyg-component', function(e) {
            if (e.keyCode === 13 && !e.shiftKey) {
                e.preventDefault();
                e.stopPropagation();

                var selection = window.getSelection();
                if (selection.rangeCount > 0) {
                    var range = selection.getRangeAt(0);

                    // Insert <br> tag
                    var br = document.createElement('br');
                    range.deleteContents();
                    range.insertNode(br);

                    // Move cursor after the <br>
                    range.setStartAfter(br);
                    range.collapse(true);

                    // If cursor is at the end, insert second <br> to ensure visible line
                    var atEnd = !range.startContainer.nextSibling ||
                               (range.startContainer.nodeType === Node.TEXT_NODE &&
                                range.startOffset >= range.startContainer.length);

                    if (atEnd) {
                        var br2 = document.createElement('br');
                        range.insertNode(br2);
                        range.setStartAfter(br2);
                        range.collapse(true);
                    }

                    selection.removeAllRanges();
                    selection.addRange(range);
                }
                return false;
            }
        });

        // Initial content will be loaded as badges after master items are fetched

        // Dropdown for master items
        var $dropdown = $('<div class="master-items-dropdown" style="position: absolute; background: white; border: 1px solid #d0d0d0; border-radius: 4px; box-shadow: 0 4px 12px rgba(0,0,0,0.15); max-height: 300px; overflow-y: auto; z-index: 10000; display: none; min-width: 300px;"></div>');

        var $searchInput = $('<input type="text" placeholder="Search master items..." style="width: 100%; padding: 10px; border: none; border-bottom: 1px solid #e0e0e0; font-size: 13px; box-sizing: border-box;">');

        var $itemsList = $('<div style="max-height: 250px; overflow-y: auto;"></div>');

        $dropdown.append($searchInput);
        $dropdown.append($itemsList);

        $editorContainer.append($editorToolbar);
        $editorContainer.append($editor);
        $editorContainer.append($dropdown);

        $body.append($info);
        $body.append($componentToolbar);
        $body.append($formatToolbar);
        $body.append($editorContainer);

        // Modal footer
        var $footer = $('<div class="modal-footer"></div>');
        var $statusMsg = $('<div class="status-message"></div>');

        var $cancelBtn = $('<button class="modal-btn cancel-btn">Cancel</button>');
        $cancelBtn.on('click', function() {
            $overlay.remove();
        });

        var $saveBtn = $('<button class="modal-btn save-btn">Save Changes</button>');
        $saveBtn.on('click', function() {
            // console.log('=== BEFORE CONVERSION ===');
            // console.log('Raw HTML:', $editor.html());
            // console.log('========================');

            // Convert badges back to code
            var content = convertBadgesToCode($editor);

            // console.log('=== AFTER CONVERSION ===');
            // console.log('Saved content:', content);
            // console.log('========================');

            if (onSave) {
                onSave(content);
            }

            $statusMsg.text('✅ Saved!').addClass('success');
            setTimeout(function() {
                $overlay.remove();
            }, 800);
        });

        // Convert badges and WYSIWYG components to code
        function convertBadgesToCode($container) {
            var $clone = $container.clone();

            // Convert WYSIWYG components back to tag syntax
            $clone.find('.wysiwyg-component').each(function() {
                var $comp = $(this);
                var componentType = $comp.attr('data-component-type');
                var attrs = $comp.attr('data-attrs') || '';

                if (componentType === 'styledbox') {
                    // Clone the component content
                    var $tempComp = $comp.clone();

                    // Remove the label div
                    $tempComp.find('[style*="position: absolute"]').remove();

                    // console.log('STYLED BOX BEFORE CONVERSION:', $tempComp.html());

                    // Convert master item badges FIRST (before text extraction)
                    $tempComp.find('.master-item-badge').each(function() {
                        var itemName = $(this).attr('data-item-name');
                        $(this).replaceWith('{{[' + itemName + ']}}');
                    });

                    // Convert <font> tags (from foreColor execCommand) to <span style>
                    $tempComp.find('font[color]').each(function() {
                        var $font = $(this);
                        var color = $font.attr('color');
                        // console.log('FOUND <font color>:', color);
                        var $span = $('<span style="color: ' + color + '"></span>');
                        $span.html($font.html());
                        $font.replaceWith($span);
                    });

                    // Convert <font> tags with size to spans
                    $tempComp.find('font[size]').each(function() {
                        var $font = $(this);
                        var size = $font.attr('size');
                        // Font size attribute mapping: 1=10px, 2=13px, 3=16px, 4=18px, 5=24px, 6=32px, 7=48px
                        var sizeMap = {'1': '10px', '2': '13px', '3': '16px', '4': '18px', '5': '24px', '6': '32px', '7': '48px'};
                        var fontSize = sizeMap[size] || '16px';
                        var $span = $('<span style="font-size: ' + fontSize + '"></span>');
                        $span.html($font.html());
                        $font.replaceWith($span);
                    });

                    // Merge nested spans FIRST before processing
                    // When we have <span style="font-size: 20px;"><span style="color: red;">Test</span></span>
                    // we need to merge them into a single span with combined styles
                    $tempComp.find('span[style]').each(function() {
                        var $outerSpan = $(this);
                        var $innerSpan = $outerSpan.children('span[style]').first();

                        if ($innerSpan.length > 0) {
                            // Merge styles from outer and inner
                            var outerStyle = $outerSpan.attr('style') || '';
                            var innerStyle = $innerSpan.attr('style') || '';

                            // Combine styles (inner takes precedence for conflicts)
                            var mergedStyle = outerStyle;
                            if (innerStyle) {
                                if (mergedStyle && !mergedStyle.endsWith(';')) mergedStyle += '; ';
                                mergedStyle += innerStyle;
                            }

                            // Create new span with merged styles and inner content
                            var innerContent = $innerSpan.html();
                            var $newSpan = $('<span style="' + mergedStyle + '">' + innerContent + '</span>');
                            $outerSpan.replaceWith($newSpan);
                        }
                    });

                    // CRITICAL: Split spans at <br> boundaries FIRST
                    // Our syntax {#color:size:text} doesn't support multiline text
                    // So we need to split each line into separate formatted blocks
                    $tempComp.find('span[style]').each(function() {
                        var $span = $(this);
                        var html = $span.html();

                        // Check if span contains <br> tags
                        if (html.indexOf('<br>') !== -1) {
                            var style = $span.attr('style');
                            // Split by <br> and create separate spans for each line
                            var lines = html.split('<br>');
                            var replacement = '';

                            for (var i = 0; i < lines.length; i++) {
                                var line = lines[i].trim();
                                if (line) {
                                    // Non-empty line - wrap in styled span
                                    replacement += '<span style="' + style + '">' + line + '</span>';
                                }
                                // Add newline between lines (except after last)
                                if (i < lines.length - 1) {
                                    replacement += '\n';
                                }
                            }

                            $span.replaceWith(replacement);
                        }
                    });

                    // Convert inline styled spans (from color/size pickers) to syntax
                    $tempComp.find('span[style]').each(function() {
                        var $span = $(this);

                        // Skip if this span contains other styled spans (will be handled by merge above)
                        if ($span.find('span[style]').length > 0) {
                            return;
                        }

                        var style = $span.attr('style') || '';
                        var text = $span.text();

                        // Extract color, font-size, font-weight from inline styles
                        var color = null;
                        var size = null;
                        var bold = false;
                        var italic = false;
                        var underline = false;

                        var colorMatch = style.match(/color:\s*([^;]+)/i);
                        if (colorMatch) {
                            color = colorMatch[1].trim();

                            // Convert RGB/RGBA to hex if needed
                            var rgbMatch = color.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*[\d.]+)?\)/i);
                            if (rgbMatch) {
                                var r = parseInt(rgbMatch[1]);
                                var g = parseInt(rgbMatch[2]);
                                var b = parseInt(rgbMatch[3]);
                                color = '#' + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
                            }
                        }

                        var sizeMatch = style.match(/font-size:\s*([^;]+)/i);
                        if (sizeMatch) size = sizeMatch[1].trim();

                        if (style.indexOf('font-weight: bold') !== -1 || style.indexOf('font-weight:bold') !== -1) {
                            bold = true;
                        }

                        if (style.indexOf('font-style: italic') !== -1 || style.indexOf('font-style:italic') !== -1) {
                            italic = true;
                        }

                        if (style.indexOf('text-decoration: underline') !== -1 || style.indexOf('text-decoration:underline') !== -1) {
                            underline = true;
                        }

                        // Check if adjacent sibling has formatting we should inherit
                        // This handles cases where the browser splits a formatted span when applying new color
                        var $prev = $span.prev('span[style]');
                        var $next = $span.next('span[style]');

                        if (!size && ($prev.length || $next.length)) {
                            var siblingStyle = $prev.length ? $prev.attr('style') : $next.attr('style');
                            if (siblingStyle) {
                                var siblingSizeMatch = siblingStyle.match(/font-size:\s*([^;]+)/i);
                                if (siblingSizeMatch && !size) {
                                    size = siblingSizeMatch[1].trim();
                                }
                                if (!bold && (siblingStyle.indexOf('font-weight: bold') !== -1 || siblingStyle.indexOf('font-weight:bold') !== -1)) {
                                    bold = true;
                                }
                            }
                        }

                        // Build the syntax based on what's available
                        var syntax = text;

                        // If we have color with size/bold, use {color:size:bold:text} format
                        if (color && (size || bold)) {
                            syntax = '{' + color;
                            if (size) syntax += ':' + size;
                            if (bold) syntax += ':bold';
                            syntax += ':' + text + '}';
                        }
                        // If only color, use {color:text} format
                        else if (color) {
                            syntax = '{' + color + ':' + text + '}';
                        }
                        // If only size (no color), use {size:value:text} format
                        else if (size) {
                            syntax = '{size:' + size + ':' + text + '}';
                        }
                        // If only bold, use **text**
                        else if (bold) {
                            syntax = '**' + text + '**';
                        }
                        // If only italic, use *text*
                        else if (italic) {
                            syntax = '*' + text + '*';
                        }
                        // If only underline, use __text__
                        else if (underline) {
                            syntax = '__' + text + '__';
                        }

                        $span.replaceWith(syntax);
                    });

                    // Convert formatted spans with data attributes to syntax
                    $tempComp.find('span[data-format-type]').each(function() {
                        var $span = $(this);
                        var type = $span.attr('data-format-type');
                        var value = $span.attr('data-format-value');
                        var text = $span.text();

                        var formatted;
                        if (type === 'style') {
                            if (value === 'bold') formatted = '**' + text + '**';
                            else if (value === 'italic') formatted = '*' + text + '*';
                            else if (value === 'underline') formatted = '__' + text + '__';
                            else formatted = text;
                        } else if (type === 'size') {
                            formatted = '{size:' + value + ':' + text + '}';
                        } else {
                            formatted = text;
                        }

                        $span.replaceWith(formatted);
                    });

                    // Remove the "STYLED BOX" label div first
                    $tempComp.find('div[contenteditable="false"]').remove();

                    // console.log('HTML before BR conversion:', $tempComp.html());

                    // Convert <br> tags to actual \n characters in a text node
                    $tempComp.find('br').each(function() {
                        $(this).replaceWith('\n');
                    });

                    // console.log('HTML after BR conversion:', $tempComp.html());

                    // Convert <div> tags to newlines (browser sometimes uses divs for paragraphs)
                    $tempComp.find('div').each(function() {
                        var $div = $(this);
                        var divContent = $div.html();
                        // Add newline before div content unless it's the first element
                        if ($div.prev().length > 0) {
                            $div.replaceWith('\n' + divContent);
                        } else {
                            $div.replaceWith(divContent);
                        }
                    });

                    // console.log('HTML after DIV conversion:', $tempComp.html());

                    // Now extract text - newlines should be preserved as text nodes
                    var content = $tempComp.text();

                    // console.log('Extracted content:', JSON.stringify(content));

                    // Reconstruct the tag: {bg:attrs}content{/bg}
                    var tag = '{bg:' + attrs + '}' + content + '{/bg}';
                    $comp.replaceWith(tag);
                }
                // Add other component types here as needed
            });

            // Convert master item badges
            $clone.find('.master-item-badge').each(function() {
                var itemName = $(this).attr('data-item-name');
                $(this).replaceWith('{{[' + itemName + ']}}');
            });

            // Convert tag badges (component tags) - for any remaining badge-style tags
            $clone.find('.tag-badge').each(function() {
                var tagText = $(this).attr('data-tag-text');
                $(this).replaceWith(tagText);
            });

            // Convert formatted spans back to syntax
            $clone.find('span[data-format-type]').each(function() {
                var $span = $(this);
                var type = $span.attr('data-format-type');
                var value = $span.attr('data-format-value');
                var text = $span.text();

                // Convert to markdown-style syntax
                var formatted;
                if (type === 'style') {
                    if (value === 'bold') formatted = '**' + text + '**';
                    else if (value === 'italic') formatted = '*' + text + '*';
                    else if (value === 'underline') formatted = '__' + text + '__';
                    else formatted = text;
                } else if (type === 'size') {
                    // Use custom syntax for size with px format
                    formatted = '{size:' + value + ':' + text + '}';
                } else {
                    formatted = text;
                }

                $span.replaceWith(formatted);
            });

            // Convert <br> and <div> tags to newlines before extracting text
            $clone.find('br').replaceWith('\n');
            $clone.find('div').each(function() {
                // Add newline before div content (except first one)
                if ($(this).prev().length > 0) {
                    $(this).before('\n');
                }
            });

            return $clone.text();
        }

        // Convert code to WYSIWYG
        function convertCodeToBadges(text) {
            // Build HTML with WYSIWYG rendering
            $editor.empty();

            // Split text into segments by tag boundaries
            var segments = [];
            var currentIndex = 0;

            // Process {bg:...}...{/bg} tags as WYSIWYG
            var bgRegex = /\{bg:([^}]+)\}([\s\S]*?)\{\/bg\}/g;
            var match;

            while ((match = bgRegex.exec(text)) !== null) {
                // Add text before this match
                if (match.index > currentIndex) {
                    segments.push({
                        type: 'text',
                        content: text.substring(currentIndex, match.index)
                    });
                }

                // Add styled box component
                var attrs = match[1];
                var content = match[2];

                // Parse attributes
                var bgColor = (attrs.match(/^#[0-9a-fA-F]{6}|^#[0-9a-fA-F]{3}/) || [])[0] || 'transparent';
                var padding = (attrs.match(/padding:([^\s]+)/) || [])[1] || '10px';
                var align = (attrs.match(/align:([^\s]+)/) || [])[1] || 'left';
                var color = (attrs.match(/color:(#[0-9a-fA-F]{6}|#[0-9a-fA-F]{3}|[a-z]+)/) || [])[1] || '#333';

                // Convert \n to <br> for WYSIWYG rendering
                content = content.replace(/\n/g, '<br>');

                // Parse inline formatting syntax: {#color:size:bold:text} or {size:value:text}
                content = content.replace(/\{([^}]+)\}/g, function(match, inner) {
                    var parts = inner.split(':');
                    if (parts.length < 2) return match; // Not a formatting tag

                    var text = parts[parts.length - 1]; // Last part is always the text
                    var styleAttr = '';
                    var isColorFirst = parts[0].match(/^#[0-9a-fA-F]{3,6}$/);

                    if (isColorFirst) {
                        // Format: {#color:size:bold:text} or {#color:text}
                        styleAttr += 'color: ' + parts[0] + ';';
                        for (var i = 1; i < parts.length - 1; i++) {
                            if (parts[i].match(/^\d+px$/)) styleAttr += ' font-size: ' + parts[i] + ';';
                            else if (parts[i] === 'bold') styleAttr += ' font-weight: bold;';
                            else if (parts[i] === 'italic') styleAttr += ' font-style: italic;';
                            else if (parts[i] === 'underline') styleAttr += ' text-decoration: underline;';
                        }
                    } else if (parts[0] === 'size' && parts.length >= 3) {
                        // Format: {size:32px:text}
                        styleAttr += 'font-size: ' + parts[1] + ';';
                    } else {
                        return match; // Not recognized format
                    }

                    return '<span style="' + styleAttr + '">' + text + '</span>';
                });

                segments.push({
                    type: 'styledbox',
                    attrs: attrs,
                    bgColor: bgColor,
                    padding: padding,
                    align: align,
                    color: color,
                    content: content
                });

                currentIndex = bgRegex.lastIndex;
            }

            // Add remaining text
            if (currentIndex < text.length) {
                segments.push({
                    type: 'text',
                    content: text.substring(currentIndex)
                });
            }

            // Now render each segment
            segments.forEach(function(segment) {
                if (segment.type === 'styledbox') {
                    // Render WYSIWYG styled box component
                    var $component = $('<div class="wysiwyg-component" data-component-type="styledbox" data-attrs="' + segment.attrs.replace(/"/g, '&quot;') + '" contenteditable="true" style="' +
                        'background-color: ' + segment.bgColor + '; ' +
                        'padding: ' + segment.padding + '; ' +
                        'text-align: ' + segment.align + ' !important; ' +
                        'color: ' + segment.color + '; ' +
                        'margin: 10px 0; ' +
                        'border-radius: 4px; ' +
                        'border: 2px dashed #999; ' +
                        'position: relative; ' +
                        'min-height: 40px;' +
                        '">' +
                        segment.content +
                        '<div contenteditable="false" style="position: absolute; top: 2px; right: 5px; font-size: 10px; color: rgba(0,0,0,0.3); pointer-events: none; font-weight: bold; user-select: none;">STYLED BOX</div>' +
                        '</div>');

                    $editor.append($component);
                } else if (segment.type === 'text') {
                    // Process text for master items and formatting
                    processTextSegment(segment.content);
                }
            });

            return; // Function appends directly to $editor
        }

        // Helper to process text segments (master items, formatting)
        function processTextSegment(text) {
            // Find all {{[Master Item Name]}} patterns
            var pattern = /\{\{(\[([^\]]+)\])\}\}/g;
            var parts = [];
            var lastIndex = 0;
            var match;

            while ((match = pattern.exec(text)) !== null) {
                // Add text before the match
                if (match.index > lastIndex) {
                    parts.push({
                        type: 'text',
                        content: text.substring(lastIndex, match.index)
                    });
                }

                // Add badge
                var itemName = match[2];
                var itemType = 'dimension'; // Default, will check against master items list

                // Check if this item exists in our master items list
                var foundItem = masterItems.find(function(item) {
                    return item.name === itemName;
                });

                if (foundItem) {
                    itemType = foundItem.type;
                }

                parts.push({
                    type: 'badge',
                    name: itemName,
                    itemType: itemType
                });

                lastIndex = pattern.lastIndex;
            }

            // Add remaining text
            if (lastIndex < text.length) {
                parts.push({
                    type: 'text',
                    content: text.substring(lastIndex)
                });
            }

            // Process parts and append to editor
            parts.forEach(function(part) {
                if (part.type === 'text') {
                    // Skip component placeholders
                    if (part.content === '___COMPONENT_PLACEHOLDER___') {
                        return;
                    }

                    // Parse text for formatting only (no component tags as badges anymore)
                    var textContent = part.content;

                    // Convert **bold** to WYSIWYG first (before italic to avoid conflicts)
                    textContent = textContent.replace(/\*\*(.+?)\*\*/g, function(match, text) {
                        var $span = $('<span data-format-type="style" data-format-value="bold" style="font-weight: bold;">' + text + '</span>');
                        return $span[0].outerHTML;
                    });

                    // Convert *italic* to WYSIWYG (only single asterisks, not part of **)
                    textContent = textContent.replace(/\b\*(.+?)\*\b/g, function(match, text) {
                        // Skip if this looks like it was part of **
                        if (match.indexOf('**') !== -1) return match;
                        var $span = $('<span data-format-type="style" data-format-value="italic" style="font-style: italic;">' + text + '</span>');
                        return $span[0].outerHTML;
                    });

                    // Convert __underline__ to WYSIWYG
                    textContent = textContent.replace(/__([^_]+)__/g, function(match, text) {
                        var $span = $('<span data-format-type="style" data-format-value="underline" style="text-decoration: underline;">' + text + '</span>');
                        return $span[0].outerHTML;
                    });

                    // Convert {size:value:text} to WYSIWYG (value is in px format like "18px")
                    textContent = textContent.replace(/\{size:(\d+px):([^}]+)\}/g, function(match, size, text) {
                        var $span = $('<span data-format-type="size" data-format-value="' + size + '" style="font-size: ' + size + ';">' + text + '</span>');
                        return $span[0].outerHTML;
                    });

                    // Append as HTML (with formatting) or text
                    if (textContent.indexOf('<span') !== -1) {
                        $editor.append($(textContent));
                    } else {
                        $editor.append(document.createTextNode(textContent));
                    }
                } else if (part.type === 'badge') {
                    var color = part.itemType === 'dimension' ? '#009845' : '#1976d2';
                    var bgColor = part.itemType === 'dimension' ? '#E8F5E9' : '#E3F2FD';
                    var icon = part.itemType === 'dimension' ? '📊' : '📈';

                    var $badge = $('<span class="master-item-badge" contenteditable="false" data-item-name="' + part.name + '" data-item-type="' + part.itemType + '" style="' +
                        'display: inline-block; ' +
                        'background: ' + bgColor + '; ' +
                        'color: ' + color + '; ' +
                        'padding: 3px 8px 3px 6px; ' +
                        'margin: 0 10px 0 2px; ' + // Increased right margin for cursor visibility
                        'border-radius: 4px; ' +
                        'font-size: 13px; ' +
                        'font-weight: 500; ' +
                        'border: 1px solid ' + color + '; ' +
                        'cursor: default; ' +
                        'white-space: nowrap; ' +
                        'font-family: -apple-system, system-ui, sans-serif;' +
                        '">' +
                        '<span style="margin-right: 4px;">' + icon + '</span>' +
                        '<span>' + part.name + '</span>' +
                        '<button class="badge-remove" style="' +
                            'margin-left: 6px; ' +
                            'background: none; ' +
                            'border: none; ' +
                            'color: ' + color + '; ' +
                            'cursor: pointer; ' +
                            'font-size: 14px; ' +
                            'font-weight: bold; ' +
                            'padding: 0; ' +
                            'line-height: 1;' +
                        '" title="Remove">&times;</button>' +
                        '</span>');

                    $badge.find('.badge-remove').on('click', function(e) {
                        e.stopPropagation();
                        $badge.remove();
                        $editor.focus();
                    });

                    $editor.append($badge);
                }
            });

            // Attach event handlers to tag badges (double-click to remove)
            $editor.find('.tag-badge').off('dblclick').on('dblclick', function() {
                $(this).remove();
                $editor.focus();
            });
        }

        $footer.append($statusMsg);
        $footer.append($cancelBtn);
        $footer.append($saveBtn);

        // Assemble modal
        $modal.append($header);
        $modal.append($body);
        $modal.append($footer);
        $overlay.append($modal);
        $('body').append($overlay);

        // Fetch master items (dimensions and measures)
        $statusMsg.text('Loading master items...').css('color', '#666');

        app.createGenericObject({
            qInfo: { qType: 'DimensionList' },
            qDimensionListDef: { qType: 'dimension', qData: { title: '/qMetaDef/title' } }
        }, function(reply) {
            if (reply.qDimensionList && reply.qDimensionList.qItems) {
                reply.qDimensionList.qItems.forEach(function(item) {
                    masterItems.push({
                        name: item.qMeta.title,
                        type: 'dimension'
                    });
                });
            }

            // Fetch measures
            app.createGenericObject({
                qInfo: { qType: 'MeasureList' },
                qMeasureListDef: { qType: 'measure', qData: { title: '/qMetaDef/title' } }
            }, function(reply) {
                if (reply.qMeasureList && reply.qMeasureList.qItems) {
                    reply.qMeasureList.qItems.forEach(function(item) {
                        masterItems.push({
                            name: item.qMeta.title,
                            type: 'measure'
                        });
                    });
                }

                // Update toolbar with counts
                $editorToolbar.html(
                    '<span style="color: #009845;">Dimensions: <strong>' + masterItems.filter(function(m) { return m.type === 'dimension'; }).length + '</strong></span> ' +
                    '<span style="margin-left: 15px; color: #1976d2;">Measures: <strong>' + masterItems.filter(function(m) { return m.type === 'measure'; }).length + '</strong></span>'
                );

                // Convert existing code to badges
                var initialText = section.markdownText || '';
                // Remove [No Data] placeholders and "No data" spans from visual editor
                initialText = initialText.replace(/\[No [Dd]ata\]/gi, '');
                initialText = initialText.replace(/<span[^>]*class="no-data"[^>]*>.*?<\/span>/gi, '');
                initialText = initialText.replace(/<span[^>]*>\[No [Dd]ata\]<\/span>/gi, '');
                // Remove horizontal rules (---) from visual editor
                initialText = initialText.replace(/^---+\s*$/gm, '');
                initialText = initialText.trim();
                if (initialText) {
                    convertCodeToBadges(initialText);
                }

                $statusMsg.text('✅ Loaded ' + masterItems.length + ' master items').css('color', '#009845');
                setTimeout(function() {
                    $statusMsg.text('');
                }, 3000);
            });
        });

        // Show dropdown function with keyboard navigation
        var selectedIndex = -1;
        var filteredItems = [];

        function showDropdown(searchTerm) {
            filteredItems = masterItems.filter(function(item) {
                return !searchTerm || item.name.toLowerCase().indexOf(searchTerm.toLowerCase()) !== -1;
            });

            selectedIndex = filteredItems.length > 0 ? 0 : -1;
            renderDropdownItems();

            // Position dropdown
            $dropdown.css({
                top: '60px',
                left: '15px'
            }).show();

            $searchInput.focus();
        }

        function renderDropdownItems() {
            $itemsList.empty();

            if (filteredItems.length === 0) {
                $itemsList.append('<div style="padding: 15px; text-align: center; color: #999;">No items found</div>');
            } else {
                filteredItems.forEach(function(item, index) {
                    var color = item.type === 'dimension' ? '#009845' : '#1976d2';
                    var icon = item.type === 'dimension' ? '📊' : '📈';
                    var isSelected = index === selectedIndex;

                    var $item = $('<div class="dropdown-item" data-index="' + index + '" data-item-name="' + item.name + '" data-item-type="' + item.type + '" style="padding: 10px 15px; cursor: pointer; border-bottom: 1px solid #f0f0f0; display: flex; align-items: center; gap: 10px; background: ' + (isSelected ? '#f5f5f5' : 'white') + ';">' +
                        '<span style="font-size: 16px;">' + icon + '</span>' +
                        '<span style="flex: 1; color: ' + color + '; font-weight: 500;">' + item.name + '</span>' +
                        '<span style="font-size: 11px; color: #999; text-transform: uppercase;">' + item.type + '</span>' +
                        '</div>');

                    // Just highlight on hover - don't re-render!
                    $item.hover(
                        function() {
                            selectedIndex = index;
                            // Just update CSS, don't re-render all items
                            $itemsList.find('.dropdown-item').css('background', 'white');
                            $(this).css('background', '#f5f5f5');
                        },
                        function() {
                            $(this).css('background', 'white');
                        }
                    );

                    $itemsList.append($item);
                });

                // Scroll selected item into view
                if (selectedIndex >= 0) {
                    var $selectedItem = $itemsList.find('.dropdown-item[data-index="' + selectedIndex + '"]');
                    if ($selectedItem.length) {
                        $selectedItem[0].scrollIntoView({ block: 'nearest' });
                    }
                }
            }
        }

        // Use event delegation - handle clicks on items or their children
        $itemsList.on('mousedown', function(e) {
            e.preventDefault();
            e.stopPropagation();

            // Find the dropdown-item (might be clicking a child span)
            var $item = $(e.target).closest('.dropdown-item');

            // Also try checking if target IS the item
            if ($item.length === 0 && $(e.target).hasClass('dropdown-item')) {
                $item = $(e.target);
            }

            if ($item.length) {
                var itemName = $item.attr('data-item-name');
                var itemType = $item.attr('data-item-type');

                if (itemName && itemType) {
                    insertMasterItem(itemName, itemType);
                    $dropdown.hide();
                    $searchInput.val('');
                }
            }
        });

        function selectCurrentItem() {
            if (selectedIndex >= 0 && selectedIndex < filteredItems.length) {
                var item = filteredItems[selectedIndex];
                insertMasterItem(item.name, item.type);
                $dropdown.hide();
                $searchInput.val('');
            }
        }

        // Save cursor position when showing dropdown
        var savedRange = null;

        // Insert master item as visual badge
        function insertMasterItem(itemName, itemType) {
            

            // Create badge element
            var color = itemType === 'dimension' ? '#009845' : '#1976d2';
            var bgColor = itemType === 'dimension' ? '#E8F5E9' : '#E3F2FD';
            var icon = itemType === 'dimension' ? '📊' : '📈';

            var $badge = $('<span class="master-item-badge" contenteditable="false" data-item-name="' + itemName + '" data-item-type="' + itemType + '" style="' +
                'display: inline-block; ' +
                'background: ' + bgColor + '; ' +
                'color: ' + color + '; ' +
                'padding: 3px 8px 3px 6px; ' +
                'margin: 0 10px 0 2px; ' + // Increased right margin for better cursor visibility
                'border-radius: 4px; ' +
                'font-size: 13px; ' +
                'font-weight: 500; ' +
                'border: 1px solid ' + color + '; ' +
                'cursor: default; ' +
                'white-space: nowrap; ' +
                'font-family: -apple-system, system-ui, sans-serif;' +
                '">' +
                '<span style="margin-right: 4px;">' + icon + '</span>' +
                '<span>' + itemName + '</span>' +
                '<button class="badge-remove" style="' +
                    'margin-left: 6px; ' +
                    'background: none; ' +
                    'border: none; ' +
                    'color: ' + color + '; ' +
                    'cursor: pointer; ' +
                    'font-size: 14px; ' +
                    'font-weight: bold; ' +
                    'padding: 0; ' +
                    'line-height: 1;' +
                '" title="Remove">&times;</button>' +
                '</span>');

            // Add remove functionality
            $badge.find('.badge-remove').on('click', function(e) {
                e.stopPropagation();
                $badge.remove();
                $editor.focus();
            });

            // Find and remove the {{ trigger while preserving existing badges
            var found = false;
            var walker = document.createTreeWalker(
                $editor[0],
                NodeFilter.SHOW_TEXT,
                null,
                false
            );

            var lastTextNode = null;
            var textNode;
            while (textNode = walker.nextNode()) {
                if (textNode.textContent.indexOf('{{') !== -1) {
                    lastTextNode = textNode;
                }
            }

            if (lastTextNode) {
                // Found the text node with {{
                var text = lastTextNode.textContent;
                var pos = text.lastIndexOf('{{');

                if (pos !== -1) {
                    // Split the text node at the {{ position
                    var beforeText = text.substring(0, pos);
                    var afterText = text.substring(pos + 2);

                    // Replace the text node content (remove {{)
                    lastTextNode.textContent = beforeText;

                    // Insert badge after this text node
                    var parent = lastTextNode.parentNode;
                    var nextSibling = lastTextNode.nextSibling;

                    if (nextSibling) {
                        parent.insertBefore($badge[0], nextSibling);
                    } else {
                        parent.appendChild($badge[0]);
                    }

                    // Add a visible spacer span after badge for cursor visibility
                    var $spacer = $('<span class="badge-spacer" contenteditable="true" style="display: inline-block; width: 8px; min-width: 8px;">&nbsp;</span>');

                    if (nextSibling) {
                        parent.insertBefore($spacer[0], nextSibling);
                    } else {
                        parent.appendChild($spacer[0]);
                    }

                    // Add text after spacer
                    if (afterText) {
                        var textNode = document.createTextNode(afterText);
                        if (nextSibling) {
                            parent.insertBefore(textNode, nextSibling);
                        } else {
                            parent.appendChild(textNode);
                        }
                    }

                    // Move cursor into the spacer (makes cursor very visible)
                    var range = document.createRange();
                    var sel = window.getSelection();
                    range.setStart($spacer[0].firstChild, 1); // Position after the nbsp
                    range.collapse(true);
                    sel.removeAllRanges();
                    sel.addRange(range);

                    found = true;
                }
            }

            if (!found) {
                // No {{ found, just append at the end
                $editor.append($badge[0]);
                $editor.append(document.createTextNode(' '));
            }

            $editor.focus();
            
        }

        // Insert component tag (e.g., #[title]..#[/title])
        function insertComponentTag(componentName) {
            var sel = window.getSelection();
            if (!sel.rangeCount) {
                $editor.focus();
                sel = window.getSelection();
            }

            var range = sel.getRangeAt(0);
            range.deleteContents();

            // Special handling for styledbox component - insert WYSIWYG
            if (componentName === 'styledbox') {
                // Create WYSIWYG styled box component with default styling
                var defaultAttrs = '#006580 padding:10px align:left color:#ffffff';
                var $boxComponent = $('<div class="wysiwyg-component" data-component-type="styledbox" data-attrs="' + defaultAttrs + '" contenteditable="true" style="' +
                    'background-color: #006580; ' +
                    'padding: 10px; ' +
                    'text-align: left !important; ' +
                    'color: #ffffff; ' +
                    'margin: 10px 0; ' +
                    'border-radius: 4px; ' +
                    'border: 2px dashed #999; ' +
                    'position: relative; ' +
                    'min-height: 40px;' + // Ensure box has height even when empty
                    '">Your content here' +
                    '<div style="position: absolute; top: 2px; right: 5px; font-size: 10px; color: rgba(255,255,255,0.5); pointer-events: none; font-weight: bold;">STYLED BOX</div>' +
                    '</div>');

                range.insertNode($boxComponent[0]);

                // Position cursor inside the box and select the placeholder text
                if ($boxComponent[0].firstChild) {
                    range.setStart($boxComponent[0].firstChild, 0);
                    range.setEnd($boxComponent[0].firstChild, $boxComponent[0].firstChild.textContent.length);
                    sel.removeAllRanges();
                    sel.addRange(range);
                }

                $editor.focus();
                return;
            }

            // For other components, insert badges (future implementation)
            var openTag = '#[' + componentName + ']';
            var closeTag = '#[/' + componentName + ']';

            var $openBadge = createTagBadge(openTag, 'component');
            var $closeBadge = createTagBadge(closeTag, 'component');

            // Add zero-width spaces for cursor positioning
            var textNodeBetween = document.createTextNode('\u200B');
            var textNodeAfter = document.createTextNode('\u200B ');

            range.insertNode(textNodeAfter);
            range.insertNode($closeBadge[0]);
            range.insertNode(textNodeBetween);
            range.insertNode($openBadge[0]);

            // Position cursor between the badges
            range.setStart(textNodeBetween, 1);
            range.collapse(true);
            sel.removeAllRanges();
            sel.addRange(range);

            $editor.focus();
        }

        // Insert formatting (size, style) - WYSIWYG
        function insertFormatting(type, value) {
            

            var sel = window.getSelection();
            if (!sel.rangeCount) {
                $editor.focus();
                return;
            }

            var range = sel.getRangeAt(0);
            var selectedText = range.toString();

            if (!selectedText) {
                alert('Please select some text first to apply formatting');
                return;
            }

            // Create span with actual formatting (WYSIWYG)
            var $span = $('<span data-format-type="' + type + '" data-format-value="' + value + '"></span>');

            // Apply visual styles
            if (type === 'size') {
                // value is already in px format (e.g., "18px")
                $span.css('font-size', value);
            } else if (type === 'style') {
                if (value === 'bold') $span.css('font-weight', 'bold');
                else if (value === 'italic') $span.css('font-style', 'italic');
                else if (value === 'underline') $span.css('text-decoration', 'underline');
            }

            $span.text(selectedText);

            // Replace selection with formatted span
            range.deleteContents();
            range.insertNode($span[0]);

            // Move cursor after span
            range.setStartAfter($span[0]);
            range.collapse(true);
            sel.removeAllRanges();
            sel.addRange(range);

            $editor.focus();
            
        }

        // Create tag badge helper
        function createTagBadge(text, type) {
            var color = type === 'component' ? '#7B1FA2' : '#E65100';
            var bgColor = type === 'component' ? '#F3E5F5' : '#FFE0B2';

            var $badge = $('<span class="tag-badge" contenteditable="false" data-tag-text="' + text + '" data-tag-type="' + type + '" style="' +
                'display: inline-block; ' +
                'background: ' + bgColor + '; ' +
                'color: ' + color + '; ' +
                'padding: 2px 6px; ' +
                'margin: 0 4px; ' + // Increased margin for better cursor visibility
                'border-radius: 3px; ' +
                'font-size: 12px; ' +
                'font-weight: 500; ' +
                'border: 1px solid ' + color + '; ' +
                'cursor: default; ' +
                'white-space: nowrap; ' +
                'font-family: monospace; ' +
                'vertical-align: baseline;' + // Align with text
                '">' + text + '</span>');

            // Double-click to remove
            $badge.on('dblclick', function() {
                $badge.remove();
                $editor.focus();
            });

            return $badge;
        }

        // Keyboard trigger removed - use toolbar button instead for better UX

        // Search in dropdown
        $searchInput.on('input', function() {
            var searchTerm = $(this).val();
            showDropdown(searchTerm);
        });

        // Keyboard navigation in dropdown
        $searchInput.on('keydown', function(e) {
            if (e.key === 'Escape') {
                $dropdown.hide();
                $editor.focus();
                e.preventDefault();
            } else if (e.key === 'ArrowDown') {
                e.preventDefault();
                if (selectedIndex < filteredItems.length - 1) {
                    selectedIndex++;
                    renderDropdownItems();
                }
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                if (selectedIndex > 0) {
                    selectedIndex--;
                    renderDropdownItems();
                }
            } else if (e.key === 'Enter') {
                e.preventDefault();
                selectCurrentItem();
            }
        });

        // Focus editor
        setTimeout(function() {
            $editor.focus();
        }, 100);
    }

    function showClaudeCoachingModal(layout, sectionsData) {
        // Collect all section data
        var goldSheetData = {
            situationalAppraisal: [],
            accountStrategy: [],
            actionPlan: []
        };

        // Group sections by their content
        sectionsData.forEach(function(sd) {
            var section = sd.section;
            var sectionText = section.label + '\n' + (section.markdownText || '');

            // Try to categorize based on label
            var label = (section.label || '').toLowerCase();
            if (label.includes('situational') || label.includes('player') || label.includes('solution') || label.includes('contract') || label.includes('competitive') || label.includes('channel') || label.includes('strength') || label.includes('vulnerab')) {
                goldSheetData.situationalAppraisal.push(sectionText);
            } else if (label.includes('strategy') || label.includes('charter') || label.includes('aspiration')) {
                goldSheetData.accountStrategy.push(sectionText);
            } else if (label.includes('action') || label.includes('plan')) {
                goldSheetData.actionPlan.push(sectionText);
            } else {
                // Default to situational appraisal
                goldSheetData.situationalAppraisal.push(sectionText);
            }
        });

        // Build formatted data string
        var formattedData = '--- GOLD SHEET DATA ---\n\n';
        formattedData += '=== SITUATIONAL APPRAISAL (Where We Are) ===\n\n';
        formattedData += goldSheetData.situationalAppraisal.join('\n\n') + '\n\n';
        formattedData += '=== ACCOUNT STRATEGY (Where We Want To Go) ===\n\n';
        formattedData += goldSheetData.accountStrategy.join('\n\n') + '\n\n';
        formattedData += '=== ACTION PLAN (How We Get There) ===\n\n';
        formattedData += goldSheetData.actionPlan.join('\n\n');

        // Build coaching prompt
        var coachingPrompt = `You are an expert sales coach specializing in the **Korn Ferry Sales Methodology** (Miller Heiman Strategic Selling). I will provide you with a completed Gold Sheet dataset for an active deal, including:

- **Situational Appraisal** – current state of the opportunity (where we are)
- **Account Strategy** – desired outcomes and positioning (where we want to go)
- **Action Plan** – planned activities and next steps (how we get there)

Based on this data, please provide a comprehensive **deal coaching report** that includes:

1. **Deal Health Assessment** – Rate the overall strength of this opportunity (High / Medium / Low confidence) and explain why, based on the data provided.

2. **Buying Influences Analysis** – Identify any gaps, red flags, or blind spots in how we are covering Economic Buyers, User Buyers, Technical Buyers, and Coaches. Call out any roles that appear uncovered, uncommitted, or at risk.

3. **Win Themes & Competitive Position** – Assess how well our value proposition and differentiation are aligned to the buyer's key business issues. Where are we strong? Where are we exposed?

4. **Key Risks & Vulnerabilities** – Identify the top 3–5 threats to winning this deal (political, competitive, relationship, timing, or solution gaps).

5. **Improvement Opportunities** – Where is the strategy weakest? What is being assumed rather than validated? What actions are missing or insufficient?

6. **Prioritized Action Plan for the Rep** – Give the sales rep a clear, prioritized list of the most important things they should do in the next 2 weeks to improve their position and advance the deal. Be specific and direct.

7. **Coach's Bottom Line** – A frank, 3–5 sentence summary of what this rep needs to hear: what's working, what's at risk, and the one thing they must focus on to win.

Be direct, specific, and constructive. Do not just summarize the data back — interpret it, challenge assumptions, and give actionable guidance a seasoned sales coach would give.

---

${formattedData}`;

        // Create modal
        var $overlay = $('<div class="edit-modal-overlay"></div>');
        var $modal = $('<div class="edit-modal" style="max-width: 900px; height: 90vh;"></div>');

        var $header = $('<div class="modal-header"></div>');
        $header.append('<h3>🤖 Claude Sales Coach</h3>');
        $header.append('<div style="font-size: 12px; color: #666; margin-top: 4px;">Korn Ferry Sales Methodology Analysis</div>');

        var $closeBtn = $('<button class="modal-close-btn">&times;</button>');
        $closeBtn.on('click', function() {
            $overlay.remove();
        });
        $header.append($closeBtn);

        var $body = $('<div class="modal-body" style="display: flex; flex-direction: column; height: calc(100% - 120px);"></div>');

        // Loading state
        var $loading = $('<div style="text-align: center; padding: 60px 20px;">' +
            '<div style="display: inline-block; width: 40px; height: 40px; border: 4px solid #f3f3f3; border-top: 4px solid #667eea; border-radius: 50%; animation: spin 1s linear infinite;"></div>' +
            '<div style="margin-top: 20px; color: #666; font-size: 16px;">Analyzing your Gold Sheet with Claude...</div>' +
            '<div style="margin-top: 8px; color: #999; font-size: 13px;">This may take 30-60 seconds</div>' +
            '</div>');

        var $result = $('<div class="markdown-content" style="flex: 1; overflow-y: auto; padding: 20px; background: #f9f9f9; border-radius: 8px; display: none;"></div>');

        $body.append($loading);
        $body.append($result);

        $modal.append($header);
        $modal.append($body);

        $overlay.append($modal);
        $('body').append($overlay);

        // Call Claude API
        var orchestratorUrl = layout.mcpOrchestratorUrl || 'https://gse-mcp.replit.app';
        var endpoint = '/api/execute-tool';
        var maxTokens = layout.claudeMaxTokens || 4000; // Longer response for coaching

        var headers = {
            'Content-Type': 'application/json'
        };

        // Build fetch options with SSO credentials
        var fetchOptions = {
            method: 'POST',
            headers: headers,
            credentials: 'include', // SSO mode - always include cookies
            body: JSON.stringify({
                serverId: 'claude-server',
                toolName: 'claude-prompt',
                parameters: {
                    prompt: coachingPrompt,
                    system_prompt: 'You are an expert sales coach specializing in the Korn Ferry Sales Methodology (Miller Heiman Strategic Selling). Provide direct, specific, and constructive coaching based on the Gold Sheet data provided.',
                    max_tokens: maxTokens
                }
            })
        };

        fetch(orchestratorUrl + endpoint, fetchOptions)
        .then(function(response) {
            if (!response.ok) {
                throw new Error('API request failed: ' + response.status);
            }
            return response.json();
        })
        .then(function(data) {
            

            $loading.hide();
            $result.show();

            // Parse response from MCP server format (same as existing Claude integration)
            var responseText = '';
            if (data.result && data.result.content && data.result.content[0] && data.result.content[0].text) {
                responseText = data.result.content[0].text;
            } else if (data.analysis) {
                responseText = data.analysis;
            } else if (data.response) {
                responseText = data.response;
            } else {
                responseText = 'No response received';
            }

            // Convert markdown to HTML
            var html = parseMarkdown(responseText);
            $result.html(html);
        })
        .catch(function(error) {
            

            $loading.hide();
            $result.show();
            $result.html('<div style="padding: 20px; background: #fee; border-left: 4px solid #c00; border-radius: 4px;">' +
                '<strong>⚠️ Error</strong><br>' +
                'Failed to get coaching analysis: ' + error.message +
                '<br><br>Please check your Claude API settings in the extension properties.' +
                '</div>');
        });
    }

    function renderSectionsUI($element, layout, sectionsData, self, instanceId) {
        sectionsData.forEach(function(sd, idx) {
            // Verbose debug logging - commented out to reduce console spam
            // 

            // CRITICAL: Clear any old DOM references that might cause conflicts
            delete sd.$element;
            delete sd.$content;
        });

        // CRITICAL: Unbind ALL click handlers from edit buttons before clearing
        // This prevents orphaned click events from triggering during re-render
        $element.find('.section-edit-btn').off('click');

        // CRITICAL: Clear the element before rendering new content
        // This ensures polling updates replace old content instead of appending
        
        $element.empty();
        

        // Add "Coach with Claude" button if user has permission
        if (userHasFeaturePermission('jz_claude')) {
            var $coachBtn = $('<button class="claude-coach-btn" title="Coach with Claude">⚛</button>');
            $coachBtn.css({
                'position': 'absolute',
                'top': '8px',
                'right': '8px',
                'z-index': '1000',
                'padding': '6px 10px',
                'background': 'white',
                'color': '#009845',
                'border': 'none',
                'border-right': '1px solid #e0e0e0',
                'border-bottom': '1px solid #e0e0e0',
                'border-radius': '4px',
                'cursor': 'pointer',
                'font-size': '20px',
                'line-height': '1',
                'box-shadow': 'none',
                'transition': 'all 0.2s ease',
                'display': 'flex',
                'align-items': 'center',
                'justify-content': 'center',
                'font-weight': 'normal'
            });

            $coachBtn.hover(
                function() {
                    $(this).css({
                        'background': 'rgba(0, 152, 69, 0.05)'
                    });
                },
                function() {
                    $(this).css({
                        'background': 'white'
                    });
                }
            );

            $coachBtn.on('click', function() {
                
                showClaudeCoachingModal(layout, sectionsData);
            });

            $element.append($coachBtn);
        }

        var $container = $('<div class="multi-section-container"></div>');
        $container.css({
            'padding-top': '0'
        });
        var spacing = layout.spacing !== undefined ? layout.spacing : 5;
        var padding = layout.padding !== undefined ? layout.padding : 3;

        var widthMap = {
            'full': 1,
            'half': 0.5,
            'third': 0.333,
            'quarter': 0.25
        };

        var columnsMap = {
            'full': 12,
            'half': 6,
            'third': 4,
            'quarter': 3
        };

        // Group sections by their parent group
        // Use groupIndex to ensure uniqueness even if labels are identical
        var groupedSections = {};
        var groupOrder = [];

        sectionsData.forEach(function(sectionData) {
            var groupLabel = sectionData.section._groupLabel || 'Default Group';
            var groupWidth = sectionData.section._groupWidth || 'full';
            var groupBgColor = sectionData.section._groupBgColor || 'transparent';
            var groupSpacing = sectionData.section._groupSpacing !== undefined ? sectionData.section._groupSpacing : spacing;
            var groupBorderColor = sectionData.section._groupBorderColor || 'transparent';
            var groupBorderWidth = sectionData.section._groupBorderWidth || 0;
            var groupBorderStyle = sectionData.section._groupBorderStyle || 'solid';
            var groupIndex = sectionData.section._groupIndex !== undefined ? sectionData.section._groupIndex : 0;

            // Create unique key using index and label
            var groupKey = 'group_' + groupIndex + '_' + groupLabel;

            if (!groupedSections[groupKey]) {
                groupedSections[groupKey] = {
                    sections: [],
                    groupWidth: groupWidth,
                    groupLabel: groupLabel,
                    groupBgColor: groupBgColor,
                    groupSpacing: groupSpacing,
                    groupBorderColor: groupBorderColor,
                    groupBorderWidth: groupBorderWidth,
                    groupBorderStyle: groupBorderStyle,
                    groupIndex: groupIndex
                };
                groupOrder.push(groupKey);
            }

            groupedSections[groupKey].sections.push(sectionData);
        });

        // Create a container for all groups with grid layout
        var $groupsContainer = $('<div class="groups-container"></div>');
        $groupsContainer.css({
            'display': 'grid',
            'grid-template-columns': 'repeat(12, 1fr)',
            'gap': spacing + 'px',
            'height': '100%',
            'grid-auto-rows': 'auto',
            'align-content': 'start',
            'padding-top': '3px'
        });
        

        // Render each group
        groupOrder.forEach(function(groupKey) {
            var groupData = groupedSections[groupKey];
            var $groupContainer = $('<div class="group-container"></div>');

            var groupGridColumns = columnsMap[groupData.groupWidth];
            var groupBgColor = groupData.groupBgColor || 'transparent';
            var groupSpacing = groupData.groupSpacing !== undefined ? groupData.groupSpacing : spacing;
            var groupBorderColor = groupData.groupBorderColor || 'transparent';
            var groupBorderWidth = groupData.groupBorderWidth || 0;
            var groupBorderStyle = groupData.groupBorderStyle || 'solid';

            // Determine if group has visual styling (background or border)
            var hasVisualStyling = groupBgColor !== 'transparent' || groupBorderWidth > 0;

            $groupContainer.css({
                'grid-column': 'span ' + groupGridColumns,
                'display': 'flex',
                'flex-direction': 'column',
                'gap': groupSpacing + 'px',
                'min-width': '0',
                'background-color': groupBgColor,
                'padding': hasVisualStyling ? padding + 'px' : '0',
                'border-radius': hasVisualStyling ? '4px' : '0',
                'border': groupBorderWidth > 0 ? groupBorderWidth + 'px ' + groupBorderStyle + ' ' + groupBorderColor : 'none'
            });

            // Group sections within this group into rows based on their widths
            var rows = [];
            var currentRow = [];
            var currentRowWidth = 0;

            groupData.sections.forEach(function(sectionData) {
            var section = sectionData.section;
            var data = sectionData.data;
            var numDimensions = sectionData.numDimensions;
            var numMeasures = sectionData.numMeasures;

            // Check if section should be hidden
            var hasData = data && data.qMatrix && data.qMatrix.length > 0;

            if (section.hideIfNoData === true && !hasData) {
                return;
            }

            var sectionWidth = section.sectionWidth || 'full';
            var width = widthMap[sectionWidth];

            // If adding this section would exceed row width, start a new row
            if (currentRowWidth > 0 && currentRowWidth + width > 1.01) {
                rows.push(currentRow);
                currentRow = [];
                currentRowWidth = 0;
            }

            currentRow.push({
                section: section,
                data: data,
                numDimensions: numDimensions,
                numMeasures: numMeasures,
                itemMapping: sectionData.itemMapping,
                notFoundItems: sectionData.notFoundItems
            });
            currentRowWidth += width;

            // If we've reached a full row, start a new one
            if (currentRowWidth >= 0.99) {
                rows.push(currentRow);
                currentRow = [];
                currentRowWidth = 0;
            }
        });

        // Add any remaining sections
        if (currentRow.length > 0) {
            rows.push(currentRow);
        }

        // Render each row
        rows.forEach(function(row) {
            var $row = $('<div class="section-row"></div>');
            $row.css({
                'display': 'grid',
                'grid-template-columns': 'repeat(12, 1fr)',
                'gap': groupSpacing + 'px'
            });

            row.forEach(function(sectionData) {
                var section = sectionData.section;
                var data = sectionData.data;
                var numDimensions = sectionData.numDimensions;
                var numMeasures = sectionData.numMeasures;

                // 

                var itemMapping = sectionData.itemMapping || {};

                // 

                var $section = $('<div class="markdown-section"></div>');
                $section.attr('data-section-style', section.sectionStyle || 'card');

                var sectionWidth = section.sectionWidth || 'full';
                var sectionBgColor = section.sectionBgColor || 'transparent';
                var gridColumns = columnsMap[sectionWidth];
                var markdownText = section.markdownText || '';


                // Check if section contains table BEFORE applying styles
                var willHaveTable = markdownText.indexOf('#[table') !== -1;

                $section.css({
                    'grid-column': 'span ' + gridColumns,
                    'padding': willHaveTable ? '0' : padding + 'px',
                    'margin-bottom': '0',
                    'background-color': sectionBgColor
                });

                if (willHaveTable) {
                }

                // Show error if master items weren't found
                if (sectionData.notFoundItems && sectionData.notFoundItems.length > 0) {
                    var errorMsg = '<div style="padding: 10px; background: #fff3cd; border-left: 4px solid #ffc107; margin-bottom: 10px; border-radius: 4px;">' +
                        '<strong>⚠️ Master items not found:</strong><br>' +
                        '<ul style="margin: 5px 0; padding-left: 20px;">';
                    sectionData.notFoundItems.forEach(function(item) {
                        errorMsg += '<li>' + item + '</li>';
                    });
                    errorMsg += '</ul>' +
                        '<small style="color: #666;">Check master items panel for exact names</small>' +
                        '</div>';
                    markdownText = errorMsg + markdownText;
                }

                // First, convert {{[Master Item Name]}} to {{dim1}}, {{measure1}} format
                // This allows all existing tag processing logic to work unchanged
                // 
                // 

                Object.keys(itemMapping).forEach(function(itemName) {
                    var mapping = itemMapping[itemName];
                    var masterPattern = '{{[' + itemName + ']}}';
                    var indexPattern;

                    if (mapping.type === 'dim') {
                        indexPattern = '{{dim' + (mapping.index + 1) + '}}';
                    } else {
                        indexPattern = '{{measure' + (mapping.index + 1) + '}}';
                    }

                    // 
                    var regex = new RegExp(masterPattern.replace(/[{}[\]]/g, '\\$&'), 'g');
                    var before = markdownText;
                    markdownText = markdownText.replace(regex, indexPattern);
                    if (before !== markdownText) {
                    } else {
                    }
                });

                // 

                // IMPORTANT: Replace placeholders with actual values BUT NOT inside ITERATING content tags
                // Iterating tags (list, table, concat, grid, kpi) need the placeholder so they can iterate through ALL rows
                // Non-iterating tags (row, title, header, box, image) should have placeholders replaced
                // This prevents color syntax {#333:{{dim1}}} from breaking while allowing #[list]{{dim1}}#[/list] to work
                if (data && data.qMatrix && data.qMatrix.length > 0) {
                    // Function to check if a match is inside an ITERATING content tag
                    function isInsideIteratingTag(text, matchIndex) {
                        // Tags that iterate through data rows and need original placeholders
                        var iteratingTags = ['list', 'table', 'concat', 'grid', 'kpi'];

                        // Look backwards from match to find if we're inside an iterating tag
                        var before = text.substring(0, matchIndex);

                        // Find the most recent opening tag before this position
                        var lastOpenMatch = null;
                        var lastOpenIndex = -1;
                        iteratingTags.forEach(function(tag) {
                            var regex = new RegExp('#\\[' + tag + '[^\\]]*\\]', 'g');
                            var match;
                            while ((match = regex.exec(before)) !== null) {
                        if (match.index > lastOpenIndex) {
                                    lastOpenIndex = match.index;
                                    lastOpenMatch = tag;
                                }
                            }
                        });

                        // If we found an opening tag, check if there's a closing tag after it
                        if (lastOpenMatch) {
                            var afterOpen = before.substring(lastOpenIndex);
                            var closeRegex = new RegExp('#\\[\\/' + lastOpenMatch + '\\]');
                            if (!closeRegex.test(afterOpen)) {
                                // We're between opening and closing tag
                        return true;
                            }
                        }

                        return false;
                    }

                    for (var dimIdx = 0; dimIdx < numDimensions; dimIdx++) {
                        var placeholder = '{{dim' + (dimIdx + 1) + '}}';
                        var regex = new RegExp(placeholder.replace(/[{}]/g, '\\$&'), 'g');
                        var cell = data.qMatrix[0][dimIdx];
                        var value = cell ? cell.qText : 'N/A';

                        // Replace only placeholders NOT inside ITERATING content tags
                        markdownText = markdownText.replace(regex, function(match, offset) {
                            return isInsideIteratingTag(markdownText, offset) ? match : value;
                        });
                    }

                    for (var meaIdx = 0; meaIdx < numMeasures; meaIdx++) {
                        var placeholder = '{{measure' + (meaIdx + 1) + '}}';
                        var regex = new RegExp(placeholder.replace(/[{}]/g, '\\$&'), 'g');
                        var cell = data.qMatrix[0][numDimensions + meaIdx];
                        var value = cell ? cell.qText : 'N/A';

                        // Replace only placeholders NOT inside ITERATING content tags
                        markdownText = markdownText.replace(regex, function(match, offset) {
                            return isInsideIteratingTag(markdownText, offset) ? match : value;
                        });
                    }
                } else {
                    // No data available - replace placeholders with helpful message
                    // Remove {{dim1}}, {{measure1}} etc. to avoid showing raw placeholders
                    markdownText = markdownText.replace(/\{\{dim\d+\}\}/g, '<span style="color: #999; font-style: italic;">[No data]</span>');
                    markdownText = markdownText.replace(/\{\{measure\d+\}\}/g, '<span style="color: #999; font-style: italic;">[No data]</span>');

                    // Also handle {{[Master Item]}} syntax that wasn't resolved
                    markdownText = markdownText.replace(/\{\{\[([^\]]+)\]\}\}/g, '<span style="color: #f57c00; font-style: italic;" title="Master item not found">[Master item &quot;$1&quot; not found]</span>');
                }

                // NOW process embedded content tags (with actual values already inserted)
                // Note: colorBy is now extracted from tag attributes, not section config
                markdownText = processContentTags(markdownText, data, numDimensions, numMeasures, layout, null, itemMapping);

                // 

                // Check if there's saved content for this section (only if multi-user enabled)
                var finalContent = markdownText;

                // IMPORTANT: Store original template HTML (before saved content) for Reset button
                var originalTemplateMarkdown = markdownText;
                var originalTemplateParsed;
                if (originalTemplateMarkdown.indexOf('<') !== -1 && originalTemplateMarkdown.indexOf('>') !== -1) {
                    originalTemplateParsed = originalTemplateMarkdown;
                } else {
                    originalTemplateParsed = parseMarkdown(originalTemplateMarkdown);
                }

                // Detect edit mode (in edit mode, always use template)
                var qlikNavigation = qlik.navigation;
                var isEditMode = false;
                try {
                    var navMode = qlikNavigation.getMode();
                    isEditMode = (navMode === 'edit');
                    
                } catch (e) {
                    // Fallback: if getMode() not available, assume not in edit mode
                    isEditMode = false;
                    
                }

                // === MULTI-USER: Only apply saved content if:
                // 1. Feature is enabled
                // 2. NOT in edit/dev mode (template changes should be visible immediately)
                // 3. useTemplate flag is not set to true
                // 4. Saved content exists
                var shouldUseSavedContent = MULTI_USER_ENABLED === 1 &&
                    !isEditMode &&
                    modifiedSectionsCache.loaded &&
                    modifiedSectionsCache.data &&
                    modifiedSectionsCache.data.sections &&
                    modifiedSectionsCache.data.sections[section.label];

                if (shouldUseSavedContent) {
                    var savedContent = modifiedSectionsCache.data.sections[section.label].content;
                    var savedModifiedBy = modifiedSectionsCache.data.sections[section.label].modifiedBy;

                    // Strip any existing label+separator from saved content (in case it was saved with older code)
                    if (section.showLabel && section.label) {
                        var baseFontSize = parseInt(layout.fontSize || '14', 10);
                        var labelSizeOffset = parseInt(layout.labelSizeOffset || '4', 10);
                        var labelSize = (baseFontSize + labelSizeOffset) + 'px';
                        var labelColor = section.labelColor || '#1a1a1a';

                        // Remove label div (check for multiple possible formats)
                        var labelPatterns = [
                            '<div style="font-size: ' + labelSize + '; color: ' + labelColor,
                            '<div style="font-size:' + labelSize + '; color:' + labelColor,
                            '<div style="font-size: ' + labelSize + '; color: ' + labelColor
                        ];

                        for (var pi = 0; pi < labelPatterns.length; pi++) {
                            if (savedContent.indexOf(labelPatterns[pi]) === 0) {
                                var labelEnd = savedContent.indexOf(section.label + '</div>');
                                if (labelEnd > -1) {
                                    savedContent = savedContent.substring(labelEnd + (section.label + '</div>').length);

                                    // Also remove separator if present (check both 1px and 2px)
                                    if (savedContent.indexOf('<div style="height: 2px;') === 0 || savedContent.indexOf('<div style="height: 1px;') === 0) {
                                        var sepEnd = savedContent.indexOf('</div>');
                                        if (sepEnd > -1) {
                                            savedContent = savedContent.substring(sepEnd + 6);
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                    }

                    // Check if saved content contains unprocessed content tags
                    if (savedContent.indexOf('#[') !== -1) {

                        savedContent = processContentTags(savedContent, data, numDimensions, numMeasures, layout, null, itemMapping);
                    }

                    finalContent = savedContent;
                } else if (isEditMode) {
                    
                }

                // === CHECK FOR SECTION CONTENT PLACEHOLDER ===
                // If saved content has ⟨⟨SECTION_CONTENT⟩⟩, replace it with freshly rendered template
                if (finalContent.indexOf('⟨⟨SECTION_CONTENT⟩⟩') > -1) {
                    // Re-render the original template with fresh data
                    var freshTemplateContent = section.markdownText || '';

                    // STEP 1: Convert {{[Master Item Name]}} to {{dim1}}, {{measure1}} format
                    Object.keys(itemMapping).forEach(function(itemName) {
                        var mapping = itemMapping[itemName];
                        var masterPattern = '{{[' + itemName + ']}}';
                        var indexPattern;

                        if (mapping.type === 'dim') {
                            indexPattern = '{{dim' + (mapping.index + 1) + '}}';
                        } else {
                            indexPattern = '{{measure' + (mapping.index + 1) + '}}';
                        }

                        var regex = new RegExp(masterPattern.replace(/[{}[\]]/g, '\\$&'), 'g');
                        freshTemplateContent = freshTemplateContent.replace(regex, indexPattern);
                    });

                    // STEP 2: Replace placeholders with actual values (outside iterating tags)
                    if (data && data.qMatrix && data.qMatrix.length > 0) {
                        function isInsideIteratingTag(text, matchIndex) {
                            var iteratingTags = ['list', 'table', 'concat', 'grid', 'kpi'];
                            var before = text.substring(0, matchIndex);
                            var lastOpenMatch = null;
                            var lastOpenIndex = -1;
                            iteratingTags.forEach(function(tag) {
                                var regex = new RegExp('#\\[' + tag + '[^\\]]*\\]', 'g');
                                var match;
                                while ((match = regex.exec(before)) !== null) {
                                    if (match.index > lastOpenIndex) {
                                        lastOpenIndex = match.index;
                                        lastOpenMatch = tag;
                                    }
                                }
                            });
                            if (lastOpenMatch) {
                                var afterOpen = before.substring(lastOpenIndex);
                                var closeRegex = new RegExp('#\\[\\/' + lastOpenMatch + '\\]');
                                if (!closeRegex.test(afterOpen)) {
                                    return true;
                                }
                            }
                            return false;
                        }

                        for (var dimIdx = 0; dimIdx < numDimensions; dimIdx++) {
                            var placeholder = '{{dim' + (dimIdx + 1) + '}}';
                            var regex = new RegExp(placeholder.replace(/[{}]/g, '\\$&'), 'g');
                            var cell = data.qMatrix[0][dimIdx];
                            var value = cell ? cell.qText : 'N/A';
                            freshTemplateContent = freshTemplateContent.replace(regex, function(match, offset) {
                                return isInsideIteratingTag(freshTemplateContent, offset) ? match : value;
                            });
                        }

                        for (var meaIdx = 0; meaIdx < numMeasures; meaIdx++) {
                            var placeholder = '{{measure' + (meaIdx + 1) + '}}';
                            var regex = new RegExp(placeholder.replace(/[{}]/g, '\\$&'), 'g');
                            var cell = data.qMatrix[0][numDimensions + meaIdx];
                            var value = cell ? cell.qText : 'N/A';
                            freshTemplateContent = freshTemplateContent.replace(regex, function(match, offset) {
                                return isInsideIteratingTag(freshTemplateContent, offset) ? match : value;
                            });
                        }
                    } else {
                        // No data - replace with [No data] message
                        freshTemplateContent = freshTemplateContent.replace(/\{\{dim\d+\}\}/g, '<span style="color: #999; font-style: italic;">[No data]</span>');
                        freshTemplateContent = freshTemplateContent.replace(/\{\{measure\d+\}\}/g, '<span style="color: #999; font-style: italic;">[No data]</span>');
                        freshTemplateContent = freshTemplateContent.replace(/\{\{\[([^\]]+)\]\}\}/g, '<span style="color: #f57c00; font-style: italic;" title="Master item not found">[Master item &quot;$1&quot; not found]</span>');
                    }

                    // STEP 3: Process content tags in the template
                    if (freshTemplateContent.indexOf('#[') !== -1) {
                        freshTemplateContent = processContentTags(freshTemplateContent, data, numDimensions, numMeasures, layout, null, itemMapping);
                    }

                    // STEP 4: Parse the template markdown
                    var freshRendered = parseMarkdown(freshTemplateContent);

                    // STEP 5: Replace the placeholder with fresh content
                    finalContent = finalContent.replace(/⟨⟨SECTION_CONTENT⟩⟩/g, freshRendered);
                }

                // Always parse markdown (safe to run on mixed HTML/markdown content)
                // This ensures plain markdown outside content tags gets parsed
                var parsedMarkdown = parseMarkdown(finalContent);

                // Prepend formatted label as title if showLabel is enabled (AFTER parsing)
                if (section.showLabel && section.label) {
                    var labelHtml = '';
                    var labelStyle = section.labelStyle || 'bold';
                    // Calculate label size: default font + offset
                    var baseFontSize = parseInt(layout.fontSize || '14', 10);
                    var labelSizeOffset = parseInt(layout.labelSizeOffset || '4', 10);
                    var labelSize = (baseFontSize + labelSizeOffset) + 'px';
                    var labelColor = section.labelColor || '#1a1a1a';
                    // Ensure boolean comparison (handle both boolean and string values from Qlik)
                    var labelSeparator = (section.labelSeparator === true || section.labelSeparator === 'true');

                    var styleAttr = 'font-size: ' + labelSize + '; color: ' + labelColor + '; margin: 0; line-height: 1.3; padding: 0 0 1px 0;';

                    if (labelStyle === 'bold') {
                        styleAttr += ' font-weight: 600;';
                    } else if (labelStyle === 'italic') {
                        styleAttr += ' font-style: italic;';
                    }

                    labelHtml = '<div style="' + styleAttr + '">' + section.label + '</div>';

                    if (labelSeparator) {
                        // Solid line with gradient, tighter spacing - use min-height for consistency
                        labelHtml += '<div style="min-height: 2px; height: 2px; background: linear-gradient(to right, ' + labelColor + ', transparent); margin-bottom: 5px; opacity: 0.5; display: block; overflow: hidden;"></div>';
                    }

                    parsedMarkdown = labelHtml + parsedMarkdown;
                    
                }

                // Always create header for icons and/or edit button
                var hasIcon = section.iconAction && section.iconAction !== 'none';
                var hasEditPermission = userHasFeaturePermission('jz_edit');
                var hasEditBtn = MULTI_USER_ENABLED === 1 && section.enableEdit && hasEditPermission;

                if (hasIcon || hasEditBtn) {
                    var $header = $('<div class="section-header-with-edit"></div>');
                    $header.css({
                        'position': 'relative',
                        'height': '0',
                        'overflow': 'visible'
                    });

                    // Add action icon if configured
                    if (hasIcon) {
                        // SVG home icon (green like edit button)
                        var homeSvg = '<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' +
                            '<path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"></path>' +
                            '<polyline points="9 22 9 12 15 12 15 22"></polyline>' +
                            '</svg>';

                        var $iconBtn = $('<button class="section-icon-btn">' + homeSvg + '</button>');

                        $iconBtn.css({
                            'position': 'absolute',
                            'top': '-8px',
                            'left': '-8px',
                            'background': 'white',
                            'border': 'none',
                            'border-right': '1px solid #e0e0e0',
                            'border-bottom': '1px solid #e0e0e0',
                            'border-radius': '4px',
                            'box-shadow': '0 1px 3px rgba(0,0,0,0.12)',
                            'cursor': 'pointer',
                            'transition': 'all 0.2s ease',
                            'padding': '8px',
                            'z-index': '10',
                            'color': '#009845',
                            'display': 'flex',
                            'align-items': 'center',
                            'justify-content': 'center'
                        });

                        // Handle icon actions
                        if (section.iconAction === 'clearSelections') {
                            $iconBtn.attr('title', 'Clear All Selections');
                            $iconBtn.on('click', function(e) {
                        e.stopPropagation();
                        var app = qlik.currApp();
                        app.clearAll();
                            });
                        } else if (section.iconAction === 'homeLink') {
                            $iconBtn.attr('title', 'Go to Home');
                            $iconBtn.on('click', function(e) {
                        e.stopPropagation();
                                // Navigate to the overview/home sheet
                        var app = qlik.currApp();
                        var global = qlik.getGlobal();
                        global.getAppList(function(list) {
                                    var currentApp = list.find(function(a) { return a.qDocId === app.id; });
                                    if (currentApp) {
                                        // Go to first sheet (typically overview)
                                        app.getList('sheet', function(items) {
                                            if (items.qAppObjectList && items.qAppObjectList.qItems.length > 0) {
                                                var sheetId = items.qAppObjectList.qItems[0].qInfo.qId;
                                                qlik.navigation.gotoSheet(sheetId);
                                            }
                                        });
                                    }
                                });
                            });
                        } else if (section.iconAction === 'customLink' && section.iconLink) {
                            $iconBtn.attr('title', 'Open Link: ' + section.iconLink);
                            $iconBtn.on('click', function(e) {
                        e.stopPropagation();
                        window.open(section.iconLink, '_blank');
                            });
                        }

                        $iconBtn.hover(
                            function() {
                                $(this).css({
                                    'color': '#007a38',
                                    'box-shadow': '0 2px 6px rgba(0,0,0,0.2)'
                                });
                            },
                            function() {
                                $(this).css({
                                    'color': '#009845',
                                    'box-shadow': '0 1px 3px rgba(0,0,0,0.12)'
                                });
                            }
                        );

                        $header.append($iconBtn);
                    }

                    // Add edit button if enabled
                    if (hasEditBtn) {
                        // SVG pencil icon
                        var pencilSvg = '<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' +
                            '<path d="M17 3a2.828 2.828 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5L17 3z"></path>' +
                            '</svg>';

                        var $editBtn = $('<button class="section-edit-btn">' + pencilSvg + '</button>');

                        // Store section info in data attributes for event delegation
                        $editBtn.data('section-label', section.label);
                        $editBtn.data('section-obj', section);
                        $editBtn.data('section-data-obj', sectionData);

                        $editBtn.css({
                            'position': 'absolute',
                            'top': '-8px',
                            'right': '-8px',
                            'background': 'white',
                            'border': 'none',
                            'border-left': '1px solid #e0e0e0',
                            'border-bottom': '1px solid #e0e0e0',
                            'border-radius': '4px',
                            'box-shadow': '0 1px 3px rgba(0,0,0,0.12)',
                            'cursor': 'pointer',
                            'opacity': '0',
                            'transition': 'opacity 0.2s ease, color 0.2s ease',
                            'padding': '8px',
                            'z-index': '10',
                            'color': '#999',
                            'display': 'flex',
                            'align-items': 'center',
                            'justify-content': 'center'
                        });

                        $editBtn.hover(
                            function() {
                                $(this).css('color', '#009845');
                            },
                            function() {
                                $(this).css('color', '#999');
                            }
                        );

                        $header.append($editBtn);

                        // Store DOM references for edit functionality
                        $editBtn.data('section-element', $section);
                        $editBtn.data('original-template-html', originalTemplateParsed);
                    }

                    var $content = $('<div class="section-content markdown-content"></div>');
                    // Add left padding if icon is present
                    if (hasIcon) {
                        $content.css('padding-left', '36px');
                    }
                    // Ensure section content uses full width for tables and other elements
                    $content.css('width', '100%');
                    // Apply default font settings from appearance
                    var defaultFontSize = layout.fontSize || '14';
                    var defaultFontFamily = layout.fontFamily || "'Source Sans Pro', sans-serif";
                    $content.css({
                        'font-size': defaultFontSize + 'px',
                        'font-family': defaultFontFamily
                    });
                    $content.html(parsedMarkdown);

                    // Check if content contains a table - if so, remove ALL padding to make table full-width
                    var hasTable = parsedMarkdown.indexOf('<table') !== -1;
                    if (hasTable) {
                        $section.css('padding', '0');
                        $content.css('padding', '0');

                        // Add padding to non-table elements (titles, text, etc.) for proper spacing
                        setTimeout(function() {
                            // Handle both direct tables and tables wrapped in divs
                            $content.find('table').parent('div').css('padding', '0').css('width', '100%');

                            // Apply padding to non-table children, but exclude label and separator
                            $content.children().not('table').not('div:has(table)').each(function() {
                                var $child = $(this);
                                var style = $child.attr('style') || '';

                                // Skip label divs (have font-size and margin pattern) and separator lines (height: 2px)
                                var isLabel = style.indexOf('font-size:') > -1 && style.indexOf('margin: 0') > -1 && style.indexOf('line-height: 1.3') > -1;
                                var isSeparator = style.indexOf('height: 2px') > -1 || style.indexOf('min-height: 2px') > -1;

                                if (!isLabel && !isSeparator) {
                                    $child.css({
                                        'padding-left': padding + 'px',
                                        'padding-right': padding + 'px'
                                    });
                                }
                            });

                            // First non-table child needs top padding (but not label or separator)
                            var $firstNonTable = $content.children().not('div:has(table)').first();
                            if ($firstNonTable.length) {
                                var style = $firstNonTable.attr('style') || '';
                                var isLabel = style.indexOf('font-size:') > -1 && style.indexOf('margin: 0') > -1;
                                var isSeparator = style.indexOf('height: 2px') > -1 || style.indexOf('min-height: 2px') > -1;

                                if (!isLabel && !isSeparator) {
                                    $firstNonTable.css('padding-top', padding + 'px');
                                }
                            }
                        }, 10);
                    }

                    // Apply pixel widths after table is rendered
                    setTimeout(function() {
                        var $table = $content.find('.jz-table-debug');
                        if ($table.length > 0) {
                            // Get actual table width in pixels
                            var tableWidthPx = $table.width();

                            // Get stored percentage widths
                            var percentages = $table.attr('data-col-percentages');
                            if (percentages) {
                                var percentArray = percentages.split(',').map(function(p) {
                                    return parseFloat(p);
                                });

                                // Calculate total overhead (padding + borders)
                                // Each cell has padding: 3px 8px = 16px horizontal padding
                                // With border-collapse, borders are shared, so roughly 2px per cell
                                var numCols = percentArray.length;
                                var paddingPerCell = 16; // 8px left + 8px right
                                var borderPerCell = 2;   // Approximate with collapse
                                var totalOverhead = numCols * (paddingPerCell + borderPerCell);

                                // Calculate content width (table width minus overhead)
                                // Subtract extra 2px safety margin to prevent overflow from rounding
                                var contentWidth = tableWidthPx - totalOverhead - 2;

                                // Calculate exact pixel widths for content using floor to prevent overflow
                                var pixelWidths = percentArray.map(function(percent) {
                                    return Math.floor((percent / 100) * contentWidth);
                                });

                                // Verify total width doesn't exceed available space
                                var totalWidth = pixelWidths.reduce(function(sum, w) { return sum + w; }, 0);
                                if (totalWidth > contentWidth) {
                                    // Reduce the widest column to compensate
                                    var maxIdx = 0;
                                    var maxWidth = pixelWidths[0];
                                    for (var i = 1; i < pixelWidths.length; i++) {
                                        if (pixelWidths[i] > maxWidth) {
                                            maxWidth = pixelWidths[i];
                                            maxIdx = i;
                                        }
                                    }
                                    pixelWidths[maxIdx] -= (totalWidth - contentWidth);
                                }

                                // Apply pixel widths directly to TH and TD elements
                                var $ths = $table.find('thead th');
                                var $trs = $table.find('tbody tr');

                                $ths.each(function(idx) {
                                    $(this).css('width', pixelWidths[idx] + 'px');
                                    $(this).css('max-width', pixelWidths[idx] + 'px');
                                    $(this).css('min-width', pixelWidths[idx] + 'px');
                                });

                                $trs.each(function() {
                                    $(this).find('td').each(function(idx) {
                                        $(this).css('width', pixelWidths[idx] + 'px');
                                        $(this).css('max-width', pixelWidths[idx] + 'px');
                                        $(this).css('min-width', pixelWidths[idx] + 'px');
                                    });
                                });


                                // Log actual rendered widths after applying pixels
                                setTimeout(function() {
                                    var thWidths = [];
                                    var thSum = 0;
                                    $ths.each(function() {
                                        var w = $(this).width();
                                        thWidths.push(w + 'px');
                                        thSum += $(this).outerWidth();
                                    });
                                }, 50);
                            }
                        }
                    }, 150);

                    $section.append($header);
                    $section.append($content);

                    // Show edit button on section hover (only if enabled)
                    if (hasEditBtn) {
                        $section.hover(
                            function() {
                                $(this).find('.section-edit-btn').css('opacity', '1');
                            },
                            function() {
                                $(this).find('.section-edit-btn').css('opacity', '0');
                            }
                        );

                        // Store content element reference for edit functionality
                        $editBtn.data('content-element', $content);
                    }
                } else {
                    // No icon or edit button - still need markdown-content wrapper for CSS
                    var $content = $('<div class="section-content markdown-content"></div>');
                    // Ensure section content uses full width for tables and other elements
                    $content.css('width', '100%');
                    // Apply default font settings from appearance
                    var defaultFontSize = layout.fontSize || '14';
                    var defaultFontFamily = layout.fontFamily || "'Source Sans Pro', sans-serif";
                    $content.css({
                        'font-size': defaultFontSize + 'px',
                        'font-family': defaultFontFamily
                    });
                    $content.html(parsedMarkdown);

                    // Check if content contains a table - if so, remove ALL padding to make table full-width
                    var hasTable = parsedMarkdown.indexOf('<table') !== -1;
                    if (hasTable) {
                        $section.css('padding', '0');
                        $content.css('padding', '0');

                        // Add padding to non-table elements (titles, text, etc.) for proper spacing
                        setTimeout(function() {
                            // Handle both direct tables and tables wrapped in divs
                            $content.find('table').parent('div').css('padding', '0').css('width', '100%');

                            // Apply padding to non-table children, but exclude label and separator
                            $content.children().not('table').not('div:has(table)').each(function() {
                                var $child = $(this);
                                var style = $child.attr('style') || '';

                                // Skip label divs (have font-size and margin pattern) and separator lines (height: 2px)
                                var isLabel = style.indexOf('font-size:') > -1 && style.indexOf('margin: 0') > -1 && style.indexOf('line-height: 1.3') > -1;
                                var isSeparator = style.indexOf('height: 2px') > -1 || style.indexOf('min-height: 2px') > -1;

                                if (!isLabel && !isSeparator) {
                                    $child.css({
                                        'padding-left': padding + 'px',
                                        'padding-right': padding + 'px'
                                    });
                                }
                            });

                            // First non-table child needs top padding (but not label or separator)
                            var $firstNonTable = $content.children().not('div:has(table)').first();
                            if ($firstNonTable.length) {
                                var style = $firstNonTable.attr('style') || '';
                                var isLabel = style.indexOf('font-size:') > -1 && style.indexOf('margin: 0') > -1;
                                var isSeparator = style.indexOf('height: 2px') > -1 || style.indexOf('min-height: 2px') > -1;

                                if (!isLabel && !isSeparator) {
                                    $firstNonTable.css('padding-top', padding + 'px');
                                }
                            }
                        }, 10);
                    }

                    // Apply pixel widths after table is rendered
                    setTimeout(function() {
                        var $table = $content.find('.jz-table-debug');
                        if ($table.length > 0) {
                            // Get actual table width in pixels
                            var tableWidthPx = $table.width();

                            // Get stored percentage widths
                            var percentages = $table.attr('data-col-percentages');
                            if (percentages) {
                                var percentArray = percentages.split(',').map(function(p) {
                                    return parseFloat(p);
                                });

                                // Calculate total overhead (padding + borders)
                                // Each cell has padding: 3px 8px = 16px horizontal padding
                                // With border-collapse, borders are shared, so roughly 2px per cell
                                var numCols = percentArray.length;
                                var paddingPerCell = 16; // 8px left + 8px right
                                var borderPerCell = 2;   // Approximate with collapse
                                var totalOverhead = numCols * (paddingPerCell + borderPerCell);

                                // Calculate content width (table width minus overhead)
                                // Subtract extra 2px safety margin to prevent overflow from rounding
                                var contentWidth = tableWidthPx - totalOverhead - 2;

                                // Calculate exact pixel widths for content using floor to prevent overflow
                                var pixelWidths = percentArray.map(function(percent) {
                                    return Math.floor((percent / 100) * contentWidth);
                                });

                                // Verify total width doesn't exceed available space
                                var totalWidth = pixelWidths.reduce(function(sum, w) { return sum + w; }, 0);
                                if (totalWidth > contentWidth) {
                                    // Reduce the widest column to compensate
                                    var maxIdx = 0;
                                    var maxWidth = pixelWidths[0];
                                    for (var i = 1; i < pixelWidths.length; i++) {
                                        if (pixelWidths[i] > maxWidth) {
                                            maxWidth = pixelWidths[i];
                                            maxIdx = i;
                                        }
                                    }
                                    pixelWidths[maxIdx] -= (totalWidth - contentWidth);
                                }

                                // Apply pixel widths directly to TH and TD elements
                                var $ths = $table.find('thead th');
                                var $trs = $table.find('tbody tr');

                                $ths.each(function(idx) {
                                    $(this).css('width', pixelWidths[idx] + 'px');
                                    $(this).css('max-width', pixelWidths[idx] + 'px');
                                    $(this).css('min-width', pixelWidths[idx] + 'px');
                                });

                                $trs.each(function() {
                                    $(this).find('td').each(function(idx) {
                                        $(this).css('width', pixelWidths[idx] + 'px');
                                        $(this).css('max-width', pixelWidths[idx] + 'px');
                                        $(this).css('min-width', pixelWidths[idx] + 'px');
                                    });
                                });


                                // Log actual rendered widths after applying pixels
                                setTimeout(function() {
                                    var thWidths = [];
                                    var thSum = 0;
                                    $ths.each(function() {
                                        var w = $(this).width();
                                        thWidths.push(w + 'px');
                                        thSum += $(this).outerWidth();
                                    });
                                }, 50);
                            }
                        }
                    }, 150);

                    $section.append($content);
                }

                $row.append($section);
            });

            $groupContainer.append($row);
        });

        // Append group container to groups container
        $groupsContainer.append($groupContainer);
        }); // End of groupOrder.forEach

        $element.append($groupsContainer);

        // === CRITICAL FIX: Wait for CSS to be loaded and applied ===
        // The extension's CSS file is loaded asynchronously by Qlik
        // We must wait for it to be parsed and applied before rendering
        var cssCheckAttempts = 0;
        var maxAttempts = 100; // 100 * 50ms = 5 seconds max

        function waitForCSS() {
            cssCheckAttempts++;

            // Create a test list item to check if CSS is applied
            var $testLI = $('<ul class="markdown-content"><li>test</li></ul>');
            $element.append($testLI);

            var $li = $testLI.find('li');
            var computedStyle = window.getComputedStyle($li[0]);
            var paddingLeft = computedStyle.getPropertyValue('padding-left');
            var position = computedStyle.getPropertyValue('position');

            $testLI.remove();


            // Check if our CSS is applied (should have padding-left: 28px and position: relative)
            if (paddingLeft === '28px' && position === 'relative') {
                continueRender();
            } else if (cssCheckAttempts >= maxAttempts) {
                continueRender();
            } else {
                // Try again in 50ms
                setTimeout(waitForCSS, 50);
            }
        }

        // Start CSS check
        waitForCSS();

        function continueRender() {
        // === LAYOUT DEBUG LOGGING ===

        var $allLists = $groupsContainer.find('ul, ol');

        // Find the first section that actually has a list
        var $sectionWithList = null;
        var $listInSection = null;
        $groupsContainer.find('.markdown-section').each(function() {
            var $section = $(this);
            var $list = $section.find('ul, ol').first();
            if ($list.length > 0) {
                $sectionWithList = $section;
                $listInSection = $list;
                return false; // break
            }
        });

        if ($sectionWithList && $listInSection) {
            var $content = $sectionWithList.find('.section-content');
            var $firstLI = $listInSection.find('li').first();

            // === CRITICAL DEBUG: Check class application ===
            // Layout debug removed


            if ($firstLI.length) {
                var liPosition = $firstLI.position();
                var liOffset = $firstLI.offset();
                var sectionOffset = $sectionWithList.offset();
                var contentOffset = $content.offset();


                // Check the ::before pseudo-element by inspecting computed styles
                var liComputedStyle = window.getComputedStyle($firstLI[0], '::before');
            }
        } else {
        }

        // Force browser reflow to ensure all absolutely positioned elements (bullets, buttons) render correctly
        // This prevents the issue where bullets appear outside sections and buttons don't work until F12 is pressed
        void $groupsContainer[0].offsetHeight;

        // Check same section after reflow
        if ($sectionWithList && $listInSection) {
            var $content = $sectionWithList.find('.section-content');
            var $firstLI = $listInSection.find('li').first();


            if ($firstLI.length) {
                var liPosition = $firstLI.position();
                var liOffset = $firstLI.offset();
                var sectionOffset = $sectionWithList.offset();
                var contentOffset = $content.offset();


                // Check ::before after reflow
                var liComputedStyle = window.getComputedStyle($firstLI[0], '::before');
            }
        }

        // Detect which groups are in the last row and make them fill vertical space
        var $groups = $groupsContainer.children('.group-container');
        if ($groups.length > 0) {
            // Calculate which groups are in the last row by tracking column positions
            var groupRows = [];
            var currentRow = [];
            var currentColumn = 0;

            $groups.each(function(index) {
                var $group = $(this);
                // Extract the span value from grid-column
                var gridColumnStyle = $group.css('grid-column');
                var gridColumnSpan = 0;

                if (gridColumnStyle.includes('span')) {
                    gridColumnSpan = parseInt(gridColumnStyle.replace('span ', ''));
                } else if (gridColumnStyle.includes('/')) {
                    var parts = gridColumnStyle.split('/');
                    gridColumnSpan = parseInt(parts[1]) - parseInt(parts[0]);
                } else {
                    gridColumnSpan = parseInt(gridColumnStyle) || 4; // Default to 4 if can't parse
                }

                // If this would exceed 12 columns, start a new row
                if (currentColumn > 0 && currentColumn + gridColumnSpan > 12) {
                    groupRows.push(currentRow);
                    currentRow = [];
                    currentColumn = 0;
                }

                currentRow.push($group);
                currentColumn += gridColumnSpan;
            });

            // Add the last row
            if (currentRow.length > 0) {
                groupRows.push(currentRow);
            }

            var numRows = groupRows.length;
            

            // Set grid-template-rows: auto for all rows except last, 1fr for last
            if (numRows > 0) {
                var rowTemplate = [];
                for (var i = 0; i < numRows - 1; i++) {
                    rowTemplate.push('auto');
                }
                rowTemplate.push('1fr');

                $groupsContainer.css({
                    'grid-template-rows': rowTemplate.join(' ')
                });

                // Mark groups in the last row
                groupRows[numRows - 1].forEach(function($group) {
                    $group.addClass('last-row-group');
                });


            }
        }

        // Force another reflow after all layout calculations to ensure proper positioning
        void $element[0].offsetHeight;

        // Check layout after second reflow
        if ($sectionWithList && $listInSection) {
            var $content = $sectionWithList.find('.section-content');
            var $firstLI = $listInSection.find('li').first();


            if ($firstLI.length) {
                var liPosition = $firstLI.position();
                var liOffset = $firstLI.offset();
                var sectionOffset = $sectionWithList.offset();
                var contentOffset = $content.offset();


                // Final check of ::before
                var liComputedStyle = window.getComputedStyle($firstLI[0], '::before');
            }
        }

        // Re-enable edit buttons now that render is complete
        var $editButtons = $element.find('.section-edit-btn');
        $editButtons.prop('disabled', false).removeClass('disabled');

        // EVENT DELEGATION: Handler is now attached ONCE in paint() function
        // This prevents multiple handlers from stacking up when this function is called multiple times

        // Load all Claude AI analysis containers
        var $claudeContainers = $element.find('.claude-analysis-container');
        if ($claudeContainers.length > 0) {
            $claudeContainers.each(function() {
                var $container = $(this);
                var prompt = $container.attr('data-prompt');
                var extractedData = {
                    rows: JSON.parse($container.attr('data-extracted') || '[]')
                };

                var orchestratorUrl = $container.attr('data-orchestrator-url');
                var endpoint = $container.attr('data-endpoint');
                var systemPrompt = $container.attr('data-system-prompt');
                var maxTokens = parseInt($container.attr('data-max-tokens'));

                // Call Claude AI analysis (SSO mode)
                callClaudeAnalysis($container, prompt, extractedData, orchestratorUrl, endpoint, systemPrompt, maxTokens);
            });
        } else {
        }

        // Load all Qlik visualizations
        var $vizContainers = $element.find('.qlik-viz-container');
        if ($vizContainers.length > 0) {
            var app = qlik.currApp();

            $vizContainers.each(function() {
                var $container = $(this);
                var vizId = $container.attr('data-viz-id');
                var vizName = $container.attr('data-viz-name');

                if (vizId) {
                    app.visualization.get(vizId).then(function(vis) {
                        vis.show($container[0], { noSelections: true });
                    }).catch(function(error) {

                        $container.html('<div style="padding: 10px; color: #d32f2f;">Failed to load visualization</div>');
                    });
                } else if (vizName) {
                    app.getObject($container[0], vizName).then(function() {
                    }).catch(function(error) {

                        $container.html('<div style="padding: 10px; color: #d32f2f;">Failed to load visualization</div>');
                    });
                }
            });
        } else {
        }

        // Clear rendering state for this instance
        if (instanceId) {
            renderingState[instanceId] = false;
            delete renderStartTime[instanceId];
            hasRenderedOnce[instanceId] = true; // Mark that this instance has rendered successfully
        }

        // Start polling for updates from other users
        startPolling(instanceId, layout, $element, sectionsData, self);


        // Add dev watermark AFTER all rendering is complete
        if (IS_DEV_VERSION) {

            // Remove any existing watermarks
            $element.find('.dev-watermark, .dev-watermark-diagonal').remove();

            // Add corner badge watermark
            var $watermark = $('<div class="dev-watermark"></div>');
            $watermark.css({
                'position': 'absolute',
                'top': '10px',
                'right': '10px',
                'background': 'rgba(255, 152, 0, 0.95)',
                'color': 'white',
                'padding': '8px 16px',
                'border-radius': '4px',
                'font-weight': 'bold',
                'font-size': '12px',
                'z-index': '9999',
                'box-shadow': '0 2px 8px rgba(0,0,0,0.3)',
                'pointer-events': 'none',
                'letter-spacing': '0.5px',
                'border': '2px solid rgba(255, 152, 0, 1)',
                'text-transform': 'uppercase'
            });
            $watermark.html('🔧 DEV VERSION');
            $element.append($watermark);

            // Add diagonal watermark overlay
            var $diagonalWatermark = $('<div class="dev-watermark-diagonal"></div>');
            $diagonalWatermark.css({
                'position': 'absolute',
                'top': '50%',
                'left': '50%',
                'transform': 'translate(-50%, -50%) rotate(-45deg)',
                'font-size': '48px',
                'font-weight': 'bold',
                'color': 'rgba(255, 152, 0, 0.15)',
                'pointer-events': 'none',
                'z-index': '1',
                'white-space': 'nowrap',
                'text-transform': 'uppercase',
                'letter-spacing': '4px',
                'user-select': 'none'
            });
            $diagonalWatermark.html('DEVELOPMENT VERSION');
            $element.append($diagonalWatermark);

        }

        } // End of continueRender()
    } // End of renderSectionsUI()

    function renderSections($element, layout, data, numDimensions, numMeasures, self) {
        var $container = $('<div class="multi-section-container"></div>');
        var spacing = layout.spacing !== undefined ? layout.spacing : 5;
        var padding = layout.padding !== undefined ? layout.padding : 3;

        // Group sections into rows based on their widths
        var rows = [];
        var currentRow = [];
        var currentRowWidth = 0;

        var widthMap = {
            'full': 1,
            'half': 0.5,
            'third': 0.333,
            'quarter': 0.25
        };

        var columnsMap = {
            'full': 12,
            'half': 6,
            'third': 4,
            'quarter': 3
        };

        // Make sure sections exist
        var sections = layout.sections || [];

        sections.forEach(function(section, idx) {
            // Only skip if hideIfNoData is explicitly TRUE and there's no data
            // Default behavior (hideIfNoData = false) should always show the section
            var hasData = data && data.qMatrix && data.qMatrix.length > 0;

            if (section.hideIfNoData === true && !hasData) {
                return;
            }

            var sectionWidth = section.sectionWidth || 'full';
            var width = widthMap[sectionWidth];

            // If adding this section would exceed row width, start a new row
            if (currentRowWidth > 0 && currentRowWidth + width > 1.01) {
                rows.push(currentRow);
                currentRow = [];
                currentRowWidth = 0;
            }

            currentRow.push(section);
            currentRowWidth += width;

            // If we've reached a full row, start a new one
            if (currentRowWidth >= 0.99) {
                rows.push(currentRow);
                currentRow = [];
                currentRowWidth = 0;
            }
        });

        // Add any remaining sections
        if (currentRow.length > 0) {
            rows.push(currentRow);
        }

        // Render each row
        rows.forEach(function(row) {
            var $row = $('<div class="section-row"></div>');
            $row.css({
                'display': 'grid',
                'grid-template-columns': 'repeat(12, 1fr)',
                'gap': spacing + 'px',
                'margin-bottom': spacing + 'px'
            });

            row.forEach(function(section) {
                var $section = $('<div class="markdown-section"></div>');
                $section.attr('data-section-style', section.sectionStyle || 'card');

                var sectionWidth = section.sectionWidth || 'full';
                var sectionBgColor = section.sectionBgColor || 'transparent';
                var gridColumns = columnsMap[sectionWidth];
                var markdownText = section.markdownText || '';


                // Check if section contains table BEFORE applying styles
                var willHaveTable = markdownText.indexOf('#[table') !== -1;

                $section.css({
                    'grid-column': 'span ' + gridColumns,
                    'padding': willHaveTable ? '0' : padding + 'px',
                    'margin-bottom': '0',
                    'background-color': sectionBgColor
                });

                if (willHaveTable) {
                }

                // Process embedded content tags first (before placeholder replacement)
                markdownText = processContentTags(markdownText, data, numDimensions, numMeasures, layout, null, {});

                // Replace dimension and measure placeholders (NO automatic list conversion)
                if (data && data.qMatrix && data.qMatrix.length > 0) {
                    // Replace dimension placeholders with simple values
                    for (var dimIdx = 0; dimIdx < numDimensions; dimIdx++) {
                        var placeholder = '{{dim' + (dimIdx + 1) + '}}';
                        var regex = new RegExp(placeholder.replace(/[{}]/g, '\\$&'), 'g');

                        // Get first value only - NO automatic list conversion
                        var cell = data.qMatrix[0][dimIdx];
                        var displayValue = cell ? cell.qText : 'N/A';
                        markdownText = markdownText.replace(regex, displayValue);
                    }

                    // Replace measure placeholders
                    for (var measureIdx = 0; measureIdx < numMeasures; measureIdx++) {
                        var measurePlaceholder = '{{measure' + (measureIdx + 1) + '}}';
                        var measureRegex = new RegExp(measurePlaceholder.replace(/[{}]/g, '\\$&'), 'g');
                        var cellIdx = numDimensions + measureIdx;

                        // If multiple rows (multiple dimension combinations), sum the measure values
                        // If single row, just show that value
                        var displayValue;
                        if (data.qMatrix.length === 1) {
                            // Single row - show the value
                            var cell = data.qMatrix[0][cellIdx];
                            displayValue = cell ? cell.qText : 'N/A';
                        } else {
                            // Multiple rows - sum numeric values, or show first if non-numeric
                            var sum = 0;
                            var hasNumeric = false;
                            for (var r = 0; r < data.qMatrix.length; r++) {
                        var cell = data.qMatrix[r][cellIdx];
                        if (cell && cell.qNum !== undefined && cell.qNum !== null && !isNaN(cell.qNum)) {
                                    sum += cell.qNum;
                                    hasNumeric = true;
                                }
                            }
                            if (hasNumeric) {
                                // Format the sum using the first cell's text format as guide
                        var firstCell = data.qMatrix[0][cellIdx];
                        if (firstCell && firstCell.qText) {
                                    // Detect decimal places from first cell
                                    var decimalMatch = firstCell.qText.match(/\.(\d+)/);
                                    var decimals = decimalMatch ? decimalMatch[1].length : 0;

                                    // Format number with detected decimal places
                                    var formattedNum;
                                    if (decimals > 0) {
                                        formattedNum = sum.toFixed(decimals).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
                                    } else {
                                        formattedNum = Math.round(sum).toLocaleString();
                                    }

                                    // Try to match the format (currency symbols, etc.)
                                    var formatMatch = firstCell.qText.match(/^([^\d-]*)(-?[\d,]+\.?\d*)(.*)$/);
                                    if (formatMatch) {
                                        var prefix = formatMatch[1] || '';
                                        var suffix = formatMatch[3] || '';
                                        displayValue = prefix + formattedNum + suffix;
                                    } else {
                                        displayValue = formattedNum;
                                    }
                                } else {
                                    displayValue = sum.toLocaleString();
                                }
                            } else {
                                // Non-numeric - show first value
                        var cell = data.qMatrix[0][cellIdx];
                        displayValue = cell ? cell.qText : 'N/A';
                            }
                        }
                        markdownText = markdownText.replace(measureRegex, displayValue);
                    }
                } else {
                    // No data - replace placeholders with example text
                    for (var i = 1; i <= 10; i++) {
                        var dimPlaceholder = '{{dim' + i + '}}';
                        var dimRegex = new RegExp(dimPlaceholder.replace(/[{}]/g, '\\$&'), 'g');
                        markdownText = markdownText.replace(dimRegex, '<em style="color: #999;">[Add dimension ' + i + ']</em>');

                        var measurePlaceholder = '{{measure' + i + '}}';
                        var measureRegex = new RegExp(measurePlaceholder.replace(/[{}]/g, '\\$&'), 'g');
                        markdownText = markdownText.replace(measureRegex, '<em style="color: #999;">[Add measure ' + i + ']</em>');
                    }
                }

                // Parse markdown with color syntax
                var html = parseMarkdown(markdownText);
                var $content = $('<div class="markdown-content"></div>');
                // Apply default font settings from appearance
                var defaultFontSize = layout.fontSize || '14';
                var defaultFontFamily = layout.fontFamily || "'Source Sans Pro', sans-serif";
                $content.css({
                    'font-size': defaultFontSize + 'px',
                    'font-family': defaultFontFamily
                });
                $content.html(html);

                $section.append($content);
                $row.append($section);
            });

            $container.append($row);
        });

        $element.append($container);

        // Load Claude AI analyses
        var $claudeContainers = $element.find('.claude-analysis-container');
        if ($claudeContainers.length > 0 && userHasFeaturePermission('jz_claude')) {
            $claudeContainers.each(function() {
                var $container = $(this);
                var prompt = $container.attr('data-prompt');
                var extractedData = JSON.parse(decodeURIComponent($container.attr('data-extracted-data')));
                var orchestratorUrl = $container.attr('data-orchestrator-url');
                var endpoint = $container.attr('data-endpoint');
                var systemPrompt = $container.attr('data-system-prompt');
                var maxTokens = parseInt($container.attr('data-max-tokens'));

                // Call Claude AI analysis (SSO mode)
                callClaudeAnalysis($container, prompt, extractedData, orchestratorUrl, endpoint, systemPrompt, maxTokens);
            });
        }

        // Load all Qlik visualizations
        var $vizContainers = $element.find('.qlik-viz-container');
        if ($vizContainers.length > 0) {
            var app = qlik.currApp();

            $vizContainers.each(function() {
                var $container = $(this);
                var vizId = $container.attr('data-viz-id');
                var vizName = $container.attr('data-viz-name');
                var containerId = $container.attr('id');

                if (!containerId) return;

                // If we have an ID, use it directly
                if (vizId) {
                    app.visualization.get(vizId).then(function(vis) {
                        vis.show(containerId);
                    }).catch(function(error) {
                        
                        $container.html('<div style="padding: 8px; text-align: center; color: #d32f2f;">Failed to load visualization: ' + vizId + '<br><small>Error: ' + error.message + '</small></div>');
                    });
                }
                // If we have a name, find the master item by name
                else if (vizName) {
                    // Get list of all master items
                    app.getList('masterobject', function(reply) {
                        var masterItem = null;

                        // Search in the app object list
                        if (reply && reply.qAppObjectList && reply.qAppObjectList.qItems) {
                            reply.qAppObjectList.qItems.forEach(function(item) {
                        if (item.qMeta && item.qMeta.title === vizName) {
                                    masterItem = item;
                                }
                            });
                        }

                        if (masterItem && masterItem.qInfo && masterItem.qInfo.qId) {
                            app.visualization.get(masterItem.qInfo.qId).then(function(vis) {
                        vis.show(containerId);
                            }).catch(function(error) {
                        
                                $container.html('<div style="padding: 8px; text-align: center; color: #d32f2f;">Failed to load visualization: ' + vizName + '<br><small>Error: ' + error.message + '</small></div>');
                            });
                        } else {
                            $container.html('<div style="padding: 8px; text-align: center; color: #d32f2f;">Visualization not found: ' + vizName + '<br><small>Make sure the name matches exactly (case-sensitive)</small></div>');
                        }
                    });
                }
            });
        }
    }

    
});
