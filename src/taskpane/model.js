/*
 *
 * =============
 * GENERAL SETUP
 * =============
 *
 */
var IS_GOOGLE = typeof UrlFetchApp != "undefined";

/*
 *
 * ==============================
 * MICROSOFT EXCEL SPECIFIC SETUP
 * ==============================
 * Please comment/uncomment this block of code for Google Sheets/Excel
 */
export { params, templateFields, tempDataStore, Data };
// ===END OF EXCEL SPECIFIC CODE===



/*
 *
 * ==========
 * DATA MODEL
 * ==========
 *
 * This holds all the information about the user and all template configuration information. 
 * Future versions of this program can add their artifact to the `templateData` object.
 *
 */

// main store for hard coded enums and metadata that define fields and artifacts
var params = {
    // enums for different artifacts
    artifactEnums: {
        requirements: 1,
        testCases: 2,
        incidents: 3,
        releases: 4,
        testRuns: 5,
        tasks: 6,
        testSteps: 7,
        testSets: 8,
        risks: 14,
        folders: 114,
        components: 99,
        users: 98,
        customLists: 97,
        customValues: 96,
    },
    dataSheetName: "database",

    // enums for different types of field - match custom field prop types where relevant
    fieldType: {
        text: 1,
        int: 2,
        num: 3,
        bool: 4,
        date: 5,
        drop: 6,
        multi: 7,
        user: 8,
        // following types don't exist as custom property types as set by Spira - but useful for defining standard field types here
        id: 109,
        subId: 110,
        component: 111, // project level field
        release: 112, // project level field
        arr: 113, // used for comma separated lists in a single cell (eg linked Ids)
        folder: 114, // don't think in reality this will be need
        //custom properties (6.13+)
        customText: 1,
        customInteger: 2,
        customDecimal: 3,
        customBoolean: 4,
        customDate: 5,
        customList: 6,
        customMultiList: 7,
        customUser: 8,
        customPassword: 9,
        customRelease: 10,
        customDateAndTime: 11,
        customAutomationHost: 12,
    },

    //enums for association between artifact types we handle in the add-in
    associationEnums: {
        req2req: 1,
        tc2req: 2,
        tc2rel: 3,
        tc2ts: 4
    },
    associationTextLabels: {
        1: "[Req. to Req.]",
        2: "[TestCase to Req.]",
        3: "[TestCase to Rel.]",
        4: "[TestCase to TestSet]",
    },

    // enums and various metadata for all artifacts potentially used by the system
    artifacts: [
        { field: 'requirements', name: 'Requirements', id: 1, hierarchical: true },
        { field: 'testCases', name: 'Test Cases', id: 2, hasFolders: true, hasSubType: true, subTypeId: 7, subTypeName: "TestSteps" },
        { field: 'incidents', name: 'Incidents', id: 3 },
        { field: 'releases', name: 'Releases', id: 4, hierarchical: true },
        { field: 'testRuns', name: 'Test Runs', id: 5, disabled: true, hidden: true },
        { field: 'tasks', name: 'Tasks', id: 6, hasFolders: true },
        { field: 'testSteps', name: 'Test Steps', id: 7, disabled: true, hidden: true, isSubType: true },
        { field: 'testSets', name: 'Test Sets', id: 8, hasFolders: true },
        { field: 'risks', name: 'Risks', id: 14 },
        { field: 'folders', name: 'Folders', id: 114, sendOnly: true, adminOnly: true, hidden: true },
        { field: 'components', name: 'Components', id: 99, adminOnly: true, noPagination: true },
        { field: 'users', name: 'Users', id: 98, disabled: false, hidden: true },
        { field: 'customLists', name: 'Custom Lists', id: 97, disabled: false, hidden: true, hasDualValues: true, hasSubType: true, subTypeId: 96, subTypeName: "customValues", skipSubCustom: true, allowsCreateOnUpdate: true, allowGetSingle: true },
        { field: 'customValues', name: 'Custom Values', id: 96, disabled: true, hidden: true, isSubType: true },
    ],
    //special cases enum
    specialCases: [
        { artifactId: 2, parameter: 'TestStepId', field: 'Description', target: "Call TC:" }
    ],
    //enum for result column
    resultColumns: {
        Requirements: 14,
        'Test Cases': 18,
    },
    //enum for parentFolder fields
    parentFolders: {
        2: "ParentTestCaseFolderId",
        6: "ParentTaskFolderId",
        8: "ParentTestSetFolderId"
    },
    //enum for FolderId fields
    IdFolders: {
        2: "TestCaseFolderId",
        6: "TaskFolderId",
        8: "TestSetFolderId"
    },
};

// each artifact has all its standard fields listed, along with important metadata - display name, field type, hard coded values set by system
var templateFields = {
    requirements: [
        { field: "RequirementId", name: "ID", type: params.fieldType.id },
        { field: "Name", name: "Name", type: params.fieldType.text, required: true, setsHierarchy: true },
        { field: "Description", name: "Description", type: params.fieldType.text },
        { field: "ReleaseId", name: "Release", type: params.fieldType.release },
        {
            field: "RequirementTypeId", name: "Type", type: params.fieldType.drop, required: true,
            bespoke: {
                url: "/requirements/types",
                idField: "RequirementTypeId",
                nameField: "Name",
                isActive: "IsActive"
            },
            // these are used in cases where the field in the override can be a calculated field. In those cases, the ID for the field is incorrect - such as here.
            displayOverride: {
                field: "RequirementTypeName",
                values: ['Epic']
            }
        },
        {
            field: "ImportanceId", name: "Importance", type: params.fieldType.drop,
            bespoke: {
                url: "/requirements/importances",
                idField: "ImportanceId",
                nameField: "Name",
                isActive: "Active"
            }
        },
        {
            field: "StatusId", name: "Status", type: params.fieldType.drop,
            bespoke: {
                url: "/requirements/statuses",
                idField: "RequirementStatusId",
                nameField: "Name",
                isActive: "Active"
            }
        },
        { field: "EstimatePoints", name: "Estimate", type: params.fieldType.num },
        { field: "AuthorId", name: "Author", type: params.fieldType.user },
        { field: "OwnerId", name: "Owner", type: params.fieldType.user },
        { field: "ComponentId", name: "Component", type: params.fieldType.component },
        { field: "CreationDate", name: "Creation Date", type: params.fieldType.date, isReadOnly: true, isHidden: true },
        { field: "Text", name: "New Comment", type: params.fieldType.text, isComment: true, isAdvanced: true },
        { field: "Association", name: "New Associated Requirement(s)", type: params.fieldType.text, isAdvanced: true, association: params.associationEnums.req2req },
        { field: "Result", name: "Advanced 'Send to Spira' Log", type: params.fieldType.text, isReadOnly: true, isComments: true, isAdvanced: true },
        { field: "ConcurrencyDate", name: "Concurrency Date", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        { field: "IndentLevel", name: "Indent Level", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        { field: "Summary", name: "Summary", type: params.fieldType.bool, isReadOnly: true, isHidden: true }
    ],

    releases: [
        { field: "ReleaseId", name: "ID", type: params.fieldType.id },
        { field: "Name", name: "Name", type: params.fieldType.text, required: true, setsHierarchy: true },
        { field: "Description", name: "Description", type: params.fieldType.text },
        { field: "VersionNumber", name: "Version Number", type: params.fieldType.text },
        {
            field: "ReleaseStatusId", name: "Status", type: params.fieldType.drop, required: true,
            values: [
                { id: 1, name: "Planned" },
                { id: 2, name: "In Progress" },
                { id: 3, name: "Completed" },
                { id: 4, name: "Closed" },
                { id: 5, name: "Deferred" },
                { id: 6, name: "Cancelled" }
            ]
        },
        {
            field: "ReleaseTypeId", name: "Type", type: params.fieldType.drop, required: true,
            values: [
                { id: 1, name: 'Major Release' },
                { id: 2, name: 'Minor Release' },
                { id: 3, name: 'Sprint' },
                { id: 4, name: 'Phase' }
            ]
        },
        { field: "CreatorId", name: "Creator", type: params.fieldType.user },
        { field: "OwnerId", name: "Owner", type: params.fieldType.user },
        { field: "StartDate", name: "Start Date", type: params.fieldType.date, required: true },
        { field: "EndDate", name: "End Date", type: params.fieldType.date, required: true },
        { field: "ResourceCount", name: "Resources", type: params.fieldType.num },
        { field: "DaysNonWorking", name: "Non Working Days", type: params.fieldType.int },
        { field: "CreationDate", name: "Creation Date", type: params.fieldType.date, isReadOnly: true, isHidden: true },
        { field: "Text", name: "New Comment", type: params.fieldType.text, isComment: true, isAdvanced: true },
        { field: "ConcurrencyDate", name: "Concurrency Date", type: params.fieldType.text, isReadOnly: true, isHidden: true },
    ],

    tasks: [
        { field: "TaskId", name: "ID", type: params.fieldType.id },
        { field: "Name", name: "Name", type: params.fieldType.text, required: true },
        { field: "Description", name: "Task Description", type: params.fieldType.text },
        { field: "ReleaseId", name: "Release", type: params.fieldType.release },
        {
            field: "TaskTypeId", name: "Type", type: params.fieldType.drop, required: true,
            bespoke: {
                url: "/tasks/types",
                idField: "TaskTypeId",
                nameField: "Name",
                isActive: "IsActive"
            }
        },
        {
            field: "TaskPriorityId", name: "Priority", type: params.fieldType.drop,
            bespoke: {
                url: "/tasks/priorities",
                idField: "PriorityId",
                nameField: "Name",
                isActive: "Active"
            }
        },
        {
            field: "TaskStatusId", name: "Status", type: params.fieldType.drop, required: true,
            bespoke: {
                url: "/tasks/statuses",
                idField: "TaskStatusId",
                nameField: "Name",
                isActive: "Active"
            }
        },
        { field: "CreatorId", name: "Creator", type: params.fieldType.user },
        { field: "OwnerId", name: "Owner", type: params.fieldType.user },
        { field: "ComponentId", name: "Component", type: params.fieldType.component, isReadOnly: true },
        {
            field: "TaskFolderId", name: "Folder", type: params.fieldType.drop, values: [],
            bespoke: {
                url: "/task-folders",
                idField: "TaskFolderId",
                nameField: "Name",
                indent: "IndentLevel",
                isProjectBased: true,
                isActive: "Active"
            }
        },
        { field: "CreationDate", name: "Creation Date", type: params.fieldType.date, isReadOnly: true, isHidden: true },
        { field: "StartDate", name: "Start Date", type: params.fieldType.date },
        { field: "EndDate", name: "End Date", type: params.fieldType.date },
        { field: "EstimatedEffort", name: "Estimated Effort (in mins)", type: params.fieldType.int },
        { field: "ActualEffort", name: "Actual Effort (in mins)", type: params.fieldType.int },
        { field: "ProjectedEffort", name: "Projected Effort (in mins)", type: params.fieldType.int, isReadOnly: true, isHidden: true },
        { field: "RemainingEffort", name: "Remaining Effort (in mins)", type: params.fieldType.int },
        { field: "RequirementId", name: "RequirementId", type: params.fieldType.int },
        { field: "ProjectId", name: "Project ID", type: params.fieldType.int, isReadOnly: true, isHidden: true },
        { field: "Text", name: "New Comment", type: params.fieldType.text, isComment: true, isAdvanced: true },
        { field: "ConcurrencyDate", name: "Concurrency Date", type: params.fieldType.text, isReadOnly: true, isHidden: true },
    ],
    testSets: [
        { field: "TestSetId", name: "ID", type: params.fieldType.id },
        { field: "Name", name: "Name", type: params.fieldType.text, required: true },
        { field: "Description", name: "Description", type: params.fieldType.text },
        { field: "ReleaseId", name: "Scheduled Release", type: params.fieldType.release },
        {
            field: "TestRunTypeId", name: "Run Type", type: params.fieldType.drop, required: true,
            values: [
                { id: 1, name: "Manual" },
                { id: 2, name: "Automated" }
            ]
        },
        {
            field: "TestSetStatusId", name: "Status", type: params.fieldType.drop, required: true,
            values: [
                { id: 1, name: "Not Started" },
                { id: 2, name: "In Progress" },
                { id: 3, name: "Completed" },
                { id: 4, name: "Blocked" },
                { id: 5, name: "Deferred" }
            ]
        },
        {
            field: "RecurrenceId", name: "Recurrence", type: params.fieldType.drop,
            values: [
                { id: 1, name: "Hourly" },
                { id: 2, name: "Daily" },
                { id: 3, name: "Weekly" },
                { id: 4, name: "Monthly" },
                { id: 5, name: "Quarterly" },
                { id: 6, name: "Yearly" }
            ]
        },
        { field: "CreatorId", name: "Creator", type: params.fieldType.user },
        { field: "OwnerId", name: "Owner", type: params.fieldType.user },
        {
            field: "TestSetFolderId", name: "Folder", type: params.fieldType.drop, values: [],
            bespoke: {
                url: "/test-set-folders",
                idField: "TestSetFolderId",
                nameField: "Name",
                indent: "IndentLevel",
                isProjectBased: true,
                isActive: "Active"
            }
        },
        { field: "CreationDate", name: "Creation Date", type: params.fieldType.date, isReadOnly: true, isHidden: true },
        { field: "PlannedDate", name: "Planned Date", type: params.fieldType.date },
        { field: "Text", name: "New Comment", type: params.fieldType.text, isComment: true, isAdvanced: true },
        { field: "ExecutionDate", name: "Execution Date", type: params.fieldType.date, isReadOnly: true, isHidden: true },
        { field: "ProjectId", name: "ProjectId", type: params.fieldType.int, isReadOnly: true, isHidden: true },
        { field: "ConcurrencyDate", name: "Concurrency Date", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        { field: "AutomationHostId", name: "AutomationHostId", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        { field: "TestConfigurationSetId", name: "TestConfigurationSetId", type: params.fieldType.text, isReadOnly: true, isHidden: true }
    ],
    incidents: [
        { field: "IncidentId", name: "Incident ID", type: params.fieldType.id },
        { field: "Name", name: "Name", type: params.fieldType.text, required: true },
        { field: "Description", name: "Description", type: params.fieldType.text, required: true },
        {
            field: "IncidentTypeId", name: "Type", type: params.fieldType.drop, values: [], required: true,
            bespoke: {
                url: "/incidents/types",
                idField: "IncidentTypeId",
                nameField: "Name",
            }
        },
        {
            field: "IncidentStatusId", name: "Status", type: params.fieldType.drop, values: [], required: true,
            bespoke: {
                url: "/incidents/statuses",
                idField: "IncidentStatusId",
                nameField: "Name",
            }
        },
        {
            field: "PriorityId", name: "Priority", type: params.fieldType.drop, values: [],
            bespoke: {
                url: "/incidents/priorities",
                idField: "PriorityId",
                nameField: "Name",
            }
        },
        {
            field: "SeverityId", name: "Severity", type: params.fieldType.drop, values: [],
            bespoke: {
                url: "/incidents/severities",
                idField: "SeverityId",
                nameField: "Name",
            }
        },
        { field: "OpenerId", name: "Detected By", type: params.fieldType.user },
        { field: "OwnerId", name: "Owner", type: params.fieldType.user },
        { field: "DetectedReleaseId", name: "Detected Release", type: params.fieldType.customRelease, showInactiveReleases: true },
        { field: "ResolvedReleaseId", name: "Planned Release", type: params.fieldType.release },
        { field: "VerifiedReleaseId", name: "Verified Release", type: params.fieldType.release },
        { field: "ComponentIds", name: "Component", type: params.fieldType.component, isMulti: true },
        { field: "CreationDate", name: "Creation Date", type: params.fieldType.date, isReadOnly: true, isHidden: true },
        { field: "StartDate", name: "Start Date", type: params.fieldType.date },
        { field: "ClosedDate", name: "Closed On", type: params.fieldType.date },
        { field: "EstimatedEffort", name: "Estimated Effort (mins)", type: params.fieldType.int },
        { field: "ActualEffort", name: "Actual Effort (mins)", type: params.fieldType.int },
        { field: "Text", name: "New Comment", type: params.fieldType.text, isComment: true, isAdvanced: true },
        { field: "ConcurrencyDate", name: "Concurrency Date", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        { field: "RemainingEffort", name: "RemainingEffort", type: params.fieldType.text, isReadOnly: true, isHidden: true }
    ],

    testCases: [
        { field: "TestCaseId", name: "Case ID", type: params.fieldType.id },
        { field: "TestStepId", name: "Step ID", type: params.fieldType.subId, isSubTypeField: true },
        { field: "Name", name: "Test Case Name", type: params.fieldType.text, required: true, blocksSubType: true },
        { field: "Description", name: "Test Case Description", type: params.fieldType.text, blocksSubType: true },
        { field: "Description", name: "Test Step Description", type: params.fieldType.text, isSubTypeField: true, requiredForSubType: true, extraDataField: "LinkedTestCaseId", extraDataPrefix: "TC" },
        { field: "Position", name: "Position", type: params.fieldType.text, isSubTypeField: true, isReadOnly: true, isHidden: true },
        { field: "ExpectedResult", name: "Test Step Expected Result", type: params.fieldType.text, isSubTypeField: true},
        { field: "SampleData", name: "Test Step Sample Data", type: params.fieldType.text, isSubTypeField: true },
        {
            field: "TestCasePriorityId", name: "Test Case Priority", type: params.fieldType.drop,
            bespoke: {
                url: "/test-cases/priorities",
                idField: "PriorityId",
                nameField: "Name",
                isActive: "Active"
            }
        },
        {
            field: "TestCaseTypeId", name: "Test Case Type", type: params.fieldType.drop, required: true,
            bespoke: {
                url: "/test-cases/types",
                idField: "TestCaseTypeId",
                nameField: "Name",
                isActive: "IsActive"
            }
        },
        {
            field: "TestCaseStatusId", name: "Test Case Status", type: params.fieldType.drop, required: true,
            bespoke: {
                url: "/test-cases/statuses",
                idField: "TestCaseStatusId",
                nameField: "Name",
                isActive: "Active"
            }
        },
        { field: "AuthorId", name: "Author", type: params.fieldType.user },
        { field: "OwnerId", name: "Test Case Owner", type: params.fieldType.user },
        { field: "ProjectId", name: "ProjectId", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        {
            field: "TestCaseFolderId", name: "Test Case Folder", type: params.fieldType.drop, values: [],
            bespoke: {
                url: "/test-folders",
                idField: "TestCaseFolderId",
                nameField: "Name",
                indent: "IndentLevel",
                isProjectBased: true,
                isActive: "Active"
            }
        },
        { field: "Requirement", name: "New Associated Requirement(s)", type: params.fieldType.text, isAdvanced: true, association: params.associationEnums.tc2req },
        { field: "Release", name: "New Associated Release(s)", type: params.fieldType.text, isAdvanced: true, association: params.associationEnums.tc2rel },
        { field: "TestSet", name: "New Associated Test Set(s)", type: params.fieldType.text, isAdvanced: true, association: params.associationEnums.tc2ts },
        { field: "Text", name: "New Comment", type: params.fieldType.text, isComment: true, isAdvanced: true },
        { field: "Result", name: "Advanced 'Send to Spira' Log", type: params.fieldType.text, isReadOnly: true, isComments: true, isAdvanced: true },
        { field: "ComponentIds", name: "Test Case Component", type: params.fieldType.component, isMulti: true },
        { field: "CreationDate", name: "Test Case Creation Date", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        { field: "ConcurrencyDate", name: "Test Case Conc. Date", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        { field: "ConcurrencyDate", name: "Test Step Conc. Date", type: params.fieldType.text, isReadOnly: true, isSubTypeField: true, isHidden: true },
        { field: "ExecutionStatusId", name: "ExecutionStatusId", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        { field: "IsSuspect", name: "IsSuspect", type: params.fieldType.bool, isReadOnly: true, isHidden: true },
        { field: "EstimatedDuration", name: "EstimatedDuration", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        { field: "AutomationEngineId", name: "AutomationEngineId", type: params.fieldType.text, isReadOnly: true, isHidden: true },
        { field: "ExecutionStatusId", name: "ExecutionStatusId", type: params.fieldType.text, isReadOnly: true, isSubTypeField: true, isHidden: true }
    ],
    folders: [
        { field: "FolderId", name: "Folder ID", type: params.fieldType.id},
        { field: "Name", name: "Name", type: params.fieldType.text, required: true, setsHierarchy: true},
        { field: "Description", name: "Description", type: params.fieldType.text }
    ],
    risks: [
        { field: "RiskId", name: "ID", type: params.fieldType.id },
        { field: "Name", name: "Name", type: params.fieldType.text, required: true },
        { field: "Description", name: "Description", type: params.fieldType.text },
        { field: "ReleaseId", name: "Release", type: params.fieldType.release },
        {
            field: "RiskTypeId", name: "Type", type: params.fieldType.drop, required: true,
            bespoke: {
                url: "/risks/types",
                idField: "RiskTypeId",
                nameField: "Name",
                isActive: "IsActive"
            }
        },
        {
            field: "RiskProbabilityId", name: "Probability", type: params.fieldType.drop,
            bespoke: {
                url: "/risks/probabilities",
                idField: "RiskProbabilityId",
                nameField: "Name",
                isActive: "Active"
            }
        },
        {
            field: "RiskImpactId", name: "Impact", type: params.fieldType.drop,
            bespoke: {
                url: "/risks/impacts",
                idField: "RiskImpactId",
                nameField: "Name",
                isActive: "Active"
            }
        },
        {
            field: "RiskStatusId", name: "Status", type: params.fieldType.drop, required: true,
            bespoke: {
                url: "/risks/statuses",
                idField: "RiskStatusId",
                nameField: "Name",
                isActive: "Active"
            }
        },
        { field: "CreatorId", name: "Creator", type: params.fieldType.user, required: true },
        { field: "OwnerId", name: "Owner", type: params.fieldType.user },
        { field: "ComponentId", name: "Component", type: params.fieldType.component },
        { field: "CreationDate", name: "Creation Date", type: params.fieldType.date, isReadOnly: true, isHidden: true },
        { field: "ClosedDate", name: "Closed Date", type: params.fieldType.date },
        { field: "ReviewDate", name: "Review Date", type: params.fieldType.date },
        { field: "RiskExposure", name: "Risk Exposure", type: params.fieldType.int, isReadOnly: true, isHidden: true },
        { field: "Text", name: "New Comment", type: params.fieldType.text, isComment: true, isAdvanced: true },
        { field: "ConcurrencyDate", name: "Concurrency Date", type: params.fieldType.text, isReadOnly: true, isHidden: true },
    ],
    components: [
        { field: "ComponentId", name: "Component ID", type: params.fieldType.id },
        { field: "Name", name: "Name", type: params.fieldType.text, required: true },
        { field: "IsActive", name: "Active?", type: params.fieldType.bool },
    ],
    users: [
        { field: "UserId", name: "User ID", type: params.fieldType.id },
        { field: "FirstName", name: "First Name", type: params.fieldType.text, required: true },
        { field: "MiddleInitial", name: "Middle Initial", type: params.fieldType.text },
        { field: "LastName", name: "Last Name", type: params.fieldType.text, required: true },
        { field: "UserName", name: "UserName", type: params.fieldType.text, required: true },
        { field: "LdapDn", name: "LDAP Distinguished Name", type: params.fieldType.text },
        { field: "EmailAddress", name: "Email Address", type: params.fieldType.text, required: true },
        { field: "Admin", name: "Admin?", type: params.fieldType.bool },
        { field: "Active", name: "Active?", type: params.fieldType.bool },
        { field: "Department", name: "Department", type: params.fieldType.text },
        { field: "password", name: "Password", type: params.fieldType.text, isHeader: true, required: true },
        { field: "password_question", name: "Password Question", type: params.fieldType.text, isHeader: true, required: true },
        { field: "password_answer", name: "password Answer", type: params.fieldType.text, isHeader: true, required: true },
        {
            field: "project_id", name: "Project ID", type: params.fieldType.drop, required: false, isHeader: true,
            bespoke: {
                url: "projects",
                idField: "ProjectId",
                nameField: "Name",
                isActive: "Active",
                isSystemWide: true
            }
        },
        {
            field: "project_role_id", name: "Project Role ID", type: params.fieldType.drop, required: false, isHeader: true,
            bespoke: {
                url: "projects-roles",
                idField: "ProjectRoleId",
                nameField: "Name",
                isActive: "Active",
                isSystemWide: true
            }
        },
    ],
    customLists: [
        { field: "CustomPropertyListId", name: "List ID", type: params.fieldType.id },
        { field: "CustomPropertyValueId", name: "Value ID", type: params.fieldType.subId, isSubTypeField: true },
        { field: "Name", name: "List Name", type: params.fieldType.text, required: true, blocksSubType: true },
        { field: "Name", name: "Value Name", type: params.fieldType.text, requiredForSubType: true, isSubTypeField: true },
        { field: "Active", name: "Active?", type: params.fieldType.bool, isTypeAndSubTypeField: true, required: true },
        { field: "SortedOnValue", name: "SortedOnValue", type: params.fieldType.text, isReadOnly: true, isHidden: true },
    ],

};

function Data() {

    this.user = {
        url: '',
        userName: '',
        api_key: '',
        roleId: 1,
        admin: false
        //TODO this is wrong and should eventually be fixed to limit what user can create or edit client side
        //when add permissions - show in some way to the user what is going on
        // maybe it's as simple as a footnote explaining why projects or artifacts are disabled
    };

    this.projects = [];

    this.templates = [];

    this.operations = [
        { name: "Add new Users to Spira", id: 1, type: "send-system", artifactId: 98 },
        { name: "Add new Artifact Folders to Spira", id: 2, type: "send-product", artifactId: 114 },
        { name: "Add new Custom Lists and Values to Spira", id: 3, type: "send-template", artifactId: 97 },
        { name: "Edit existing Custom Lists and Values from Spira", id: 4, type: "get-template", artifactId: 97 },

    ];

    this.artifactFolders = [
        { name: "Test Cases", id: 2, field: "folders", mainArtifactId: 114, hierarchical: true  },
        { name: "Test Sets", id: 8, field: "folders", mainArtifactId: 114, hierarchical: true  },
        { name: "Tasks", id: 6, field: "folders", mainArtifactId: 114, hierarchical: true  },
    ];

    this.templateLists = [];

    this.currentProject = '';
    this.currentTemplate = '';
    this.projectComponents = [];
    this.projectActiveReleases = [];
    this.projectReleases = [];
    this.projectUsers = [];
    this.indentCharacter = ">";

    this.currentArtifact = '';

    this.currentOperation = '';
    this.currentList = '';

    this.projectGetRequestsToMake = 3; // users, components, releases
    this.projectGetRequestsMade = 0;

    // counts of artifact specific calls to make
    this.baselineArtifactGetRequests = 1;
    this.artifactGetRequestsToMake = this.baselineArtifactGetRequests;
    this.artifactGetRequestsMade = 0;


    this.artifactData = '';

    this.colors = {
        bgHeader: '#f1a42b',
        bgHeaderSubType: '#fdcb26',
        bgHeaderTypeAndSubType: '#ffff00',
        bgReadOnly: '#eeeeee',
        header: '#ffffff',
        headerRequired: '#000000',
        warning: '#ffcccc'
    };

    this.isTemplateLoaded = false;
    this.isGettingDataAttempt = false;
    this.fields = [];
}

function tempDataStore() {
    this.currentProject = '';
    this.currentTemplate = '';
    this.projectComponents = [];
    this.projectActiveReleases = [];
    this.projectReleases = [];
    this.projectUsers = [];

    this.currentArtifact = '';
    this.artifactCustomFields = [];
}