/*
 *
 * ==============================
 * MICROSOFT EXCEL SPECIFIC SETUP
 * ==============================
 *
 */
export { params, templateFields, tempDataStore, Data };

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
        testSets: 8
    },

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
        id: 9,
        subId: 10,
        component: 11, // project level field
        release: 12, // project level field
        arr: 13, // used for comma separated lists in a single cell (eg linked Ids)
        folder: 14 // don't think in reality this will be need
    },

    // enums and various metadata for all artifacts potentially used by the system
    artifacts: [
        {field: 'requirements', name: 'Requirements', id: 1, hierarchical: true},
        {field: 'testCases',    name: 'Test Cases',   id: 2, hasFolders: true, hasSubType: true, subTypeId: 7, subTypeName: "TestSteps"},
        {field: 'incidents',    name: 'Incidents',    id: 3},
        {field: 'releases',     name: 'Releases',     id: 4, hierarchical: true},
        {field: 'testRuns',     name: 'Test Runs',    id: 5, disabled: true, hidden: true},
        {field: 'tasks',        name: 'Tasks',        id: 6, hasFolders: true},
        {field: 'testSteps',    name: 'Test Steps',   id: 7, disabled: true, hidden: true},
        {field: 'testSets',     name: 'Test Sets',    id: 8, hasFolders: true, disabled: true, hidden: true}
    ]
};

// each artifact has all its standard fields listed, along with important metadata - display name, field type, hard coded values set by system
var templateFields = {
    requirements: [
        {field: "RequirementId", name: "ID", type: params.fieldType.id},
        {field: "Name", name: "Name", type: params.fieldType.text, required: true, setsHierarchy: true},
        {field: "Description", name: "Description", type: params.fieldType.text},
        {field: "ReleaseId", name: "Release", type: params.fieldType.release},
        {field: "RequirementTypeId", name: "Type", type: params.fieldType.drop, required: true, 
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
        {field: "ImportanceId", name: "Importance", type: params.fieldType.drop, 
            bespoke: {
                url: "/requirements/importances", 
                idField: "ImportanceId", 
                nameField: "Name", 
                isActive: "Active"
            }
        },
        {field: "StatusId", name: "Status", type: params.fieldType.drop, 
            bespoke: {
                url: "/requirements/statuses", 
                idField: "RequirementStatusId", 
                nameField: "Name", 
                isActive: "Active"
            }
        },
        {field: "EstimatePoints", name: "Estimate", type: params.fieldType.num},
        {field: "AuthorId", name: "Author", type: params.fieldType.user},
        {field: "OwnerId", name: "Owner", type: params.fieldType.user},
        {field: "ComponentId", name: "Component", type: params.fieldType.component},
        {field: "CreationDate", name: "Creation Date", type: params.fieldType.date},
    ],

    releases: [
        {field: "ReleaseId", name: "ID", type: params.fieldType.id},
        {field: "Name", name: "Name", type: params.fieldType.text, required: true, setsHierarchy: true},
        {field: "Description", name: "Description", type: params.fieldType.text},
        {field: "VersionNumber", name: "Version Number", type: params.fieldType.text},
        {field: "ReleaseStatusId", name: "Status", type: params.fieldType.drop, required: true, 
            values: [
                {id: 1, name: "Planned"},
                {id: 2, name: "In Progress"} ,
                {id: 3, name: "Completed"},
                {id: 4, name: "Closed"},
                {id: 5, name: "Deferred"},
                {id: 6, name: "Cancelled"}
            ]
        },
        {field: "ReleaseTypeId", name: "Type", type: params.fieldType.drop, required: true,
            values: [
                {id: 1, name: 'Major Release'},
                {id: 2, name: 'Minor Release'},
                {id: 3, name: 'Sprint'},
                {id: 4, name: 'Phase'}
            ]
        },
        {field: "CreatorId", name: "Creator", type: params.fieldType.user},
        {field: "OwnerId", name: "Owner", type: params.fieldType.user},
        {field: "StartDate", name: "Start Date", type: params.fieldType.date, required: true},
        {field: "EndDate", name: "End Date", type: params.fieldType.date, required: true},
        {field: "ResourceCount", name: "Resources", type: params.fieldType.num},
        {field: "DaysNonWorking", name: "Non Working Days", type: params.fieldType.int},
        {field: "CreationDate", name: "Creation Date", type: params.fieldType.date},
        // unsupported {field: "Comment", name: "Comment", type: params.fieldType.text, unsuppored: true}
    ],

    tasks: [
        {field: "TaskId", name: "ID", type: params.fieldType.id},
        {field: "Name", name: "Name", type: params.fieldType.text, required: true},
        {field: "Description", name: "Task Description", type: params.fieldType.text},
        {field: "ReleaseId", name: "Release", type: params.fieldType.release},
        {field: "TaskTypeId", name: "Type", type: params.fieldType.drop, required: true, 
            bespoke: {
                url: "/tasks/types", 
                idField: "TaskTypeId", 
                nameField: "Name", 
                isActive: "IsActive"
            }
        },
        {field: "TaskPriorityId", name: "Priority", type: params.fieldType.drop,
            bespoke: {
                url: "/tasks/priorities", 
                idField: "PriorityId", 
                nameField: "Name", 
                isActive: "Active"
            }
        },
        {field: "TaskStatusId", name: "Status", type: params.fieldType.drop, required: true,
            bespoke: {
                url: "/tasks/statuses", 
                idField: "TaskStatusId", 
                nameField: "Name", 
                isActive: "Active"
            }
        },
        {field: "CreatorId", name: "Creator", type: params.fieldType.user},
        {field: "OwnerId", name: "Owner", type: params.fieldType.user},
        {field: "ComponentId", name: "Component", type: params.fieldType.component, isReadOnly: true},
        {field: "TaskFolderId", name: "Folder", type: params.fieldType.drop, values: [],
            bespoke: {
                url: "/task-folders", 
                idField: "TaskFolderId", 
                nameField: "Name", 
                indent: "IndentLevel",
                isProjectBased: true
            }
        },  
        {field: "CreationDate", name: "Creation Date", type: params.fieldType.date},            
        {field: "StartDate", name: "Start Date", type: params.fieldType.date},
        {field: "EndDate", name: "End Date", type: params.fieldType.date},
        {field: "EstimatedEffort", name: "Estimated Effort (in mins)", type: params.fieldType.int},
        {field: "ActualEffort", name: "Actual Effort (in mins)", type: params.fieldType.int}, 
        {field: "RemainingEffort", name: "Remaining Effort (in mins)", type: params.fieldType.int}, 
    ],

    incidents: [
        {field: "IncidentId", name: "Incident ID", type: params.fieldType.id},
        {field: "Name", name: "Name", type: params.fieldType.text, required: true},
        {field: "Description", name: "Description", type: params.fieldType.text, required: true},
        {field: "IncidentTypeId", name: "Type", type: params.fieldType.drop, values: [], required: true, 
            bespoke: {
                url: "/incidents/types", 
                idField: "IncidentTypeId", 
                nameField: "Name", 
            }
        },
        {field: "IncidentStatusId", name: "Status", type: params.fieldType.drop, values: [], required: true, 
            bespoke: {
                url: "/incidents/statuses", 
                idField: "IncidentStatusId", 
                nameField: "Name", 
            }
        },
        {field: "SeverityId", name: "Severity", type: params.fieldType.drop, values: [], 
            bespoke: {
                url: "/incidents/severities", 
                idField: "SeverityId", 
                nameField: "Name", 
            }
        },
        {field: "PriorityId", name: "Priority", type: params.fieldType.drop, values: [], 
            bespoke: {
                url: "/incidents/priorities", 
                idField: "PriorityId", 
                nameField: "Name", 
            }
        },
        {field: "OpenerId", name: "Detected By", type: params.fieldType.user},
        {field: "OwnerId", name: "Owner", type: params.fieldType.user},
        {field: "DetectedReleaseId", name: "Detected Release", type: params.fieldType.release},
        {field: "ResolvedReleaseId", name: "Planned Release", type: params.fieldType.release},
        {field: "VerifiedReleaseId", name: "Verified Release", type: params.fieldType.release},
        {field: "ComponentIds", name: "Component", type: params.fieldType.component, isMulti: true},   
        {field: "CreationDate", name: "Date Detected", type: params.fieldType.date},
        {field: "StartDate", name: "Start Date", type: params.fieldType.date},
        {field: "ClosedDate", name: "Closed On", type: params.fieldType.date},
        {field: "EstimatedEffort", name: "Estimated Effort (mins)", type: params.fieldType.int},
        {field: "ActualEffort", name: "Actual Effort (mins)", type: params.fieldType.int},
        {field: "RemainingEffort", name: "Remaining Effort (mins)", type: params.fieldType.int}
    ],
    
    testCases: [
        {field: "TestCaseId", name: "Case ID", type: params.fieldType.id},
        {field: "TestStepId", name: "Step ID", type: params.fieldType.subId, isSubTypeField: true},
        {field: "Name", name: "Test Case Name", type: params.fieldType.text, required: true, blocksSubType: true},
        {field: "Description", name: "Test Case Description", type: params.fieldType.text, blocksSubType: true},
        {field: "Description", name: "Test Step Description", type: params.fieldType.text, isSubTypeField: true, requiredForSubType: true, extraDataField: "LinkedTestCaseId", extraDataPrefix: "TC"},
        {field: "ExpectedResult", name: "Test Step Expected Result", type: params.fieldType.text, isSubTypeField: true, requiredForSubType: true},
        {field: "SampleData", name: "Test Step Sample Data", type: params.fieldType.text, isSubTypeField: true},
        {field: "TestCasePriorityId", name: "Test Case Priority", type: params.fieldType.drop, 
            bespoke: {
                url: "/test-cases/priorities", 
                idField: "PriorityId", 
                nameField: "Name", 
                isActive: "Active"
            }
        },
        {field: "TestCaseTypeId", name: "Test Case Type", type: params.fieldType.drop, required: true, 
            bespoke: {
                url: "/test-cases/types", 
                idField: "TestCaseTypeId", 
                nameField: "Name", 
                isActive: "IsActive"
            }
        },
        {field: "TestCaseStatusId", name: "Test Case Status", type: params.fieldType.drop, required: true,
            bespoke: {
                url: "/test-cases/statuses", 
                idField: "TestCaseStatusId", 
                nameField: "Name", 
                isActive: "Active"
            }
        },
        {field: "OwnerId", name: "Test Case Owner", type: params.fieldType.user},
        {field: "TestCaseFolderId", name: "Test Case Folder", type: params.fieldType.drop, values: [], 
            bespoke: {
              url: "/test-folders", 
              idField: "TestCaseFolderId", 
              nameField: "Name", 
              indent: "IndentLevel",
              isProjectBased: true
            }
        },
        {field: "ComponentIds", name: "Test Case Component", type: params.fieldType.component, isMulti: true}
     ]
};

function Data() {

    this.user = {
        url: '',
        userName: '',
        api_key: '',
        roleId: 1, 
        //TODO this is wrong and should eventually be fixed to limit what user can create or edit client side
        //when add permissions - show in some way to the user what is going on
        // maybe it's as simple as a footnote explaining why projects or artifacts are disabled
    };

    this.projects = [];

    this.currentProject = '';
    this.projectComponents = [];
    this.projectReleases = [];
    this.projectUsers = [];
    this.indentCharacter = ">";
    
    this.currentArtifact = '';

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
    this.projectComponents = [];
    this.projectReleases = [];
    this.projectUsers = [];
    
    this.currentArtifact = '';
    this.artifactCustomFields = [];
}