;--------------------------------------------------------------------
; Microsoft ADO
;
; Copyright (c) 1996-1998 Microsoft Corporation.
;
;
;
; ADO constants include file for VBScript
; 
;
; convert to PB by Kiffi
;
;--------------------------------------------------------------------

;---- CursorTypeEnum Values ----
#adOpenForwardOnly = 0
#adOpenKeyset = 1
#adOpenDynamic = 2
#adOpenStatic = 3

;---- CursorOptionEnum Values ----
#adHoldRecords = $00000100
#adMovePrevious = $00000200
#adAddNew = $01000400
#adDelete = $01000800
#adUpdate = $01008000
#adBookmark = $00002000
#adApproxPosition = $00004000
#adUpdateBatch = $00010000
#adResync = $00020000
#adNotify = $00040000
#adFind = $00080000
#adSeek = $00400000
#adIndex = $00800000

;---- LockTypeEnum Values ----
#adLockReadOnly = 1
#adLockPessimistic = 2
#adLockOptimistic = 3
#adLockBatchOptimistic = 4

;---- ExecuteOptionEnum Values ----
#adAsyncExecute = $00000010
#adAsyncFetch = $00000020
#adAsyncFetchNonBlocking = $00000040
#adExecuteNoRecords = $00000080
#adExecuteStream = $00000400

;---- ConnectOptionEnum Values ----
#adAsyncConnect = $00000010

;---- ObjectStateEnum Values ----
#adStateClosed = $00000000
#adStateOpen = $00000001
#adStateConnecting = $00000002
#adStateExecuting = $00000004
#adStateFetching = $00000008

;---- CursorLocationEnum Values ----
#adUseServer = 2
#adUseClient = 3

;---- DataTypeEnum Values ----
#adEmpty = 0
#adTinyInt = 16
#adSmallInt = 2
#adInteger = 3
#adBigInt = 20
#adUnsignedTinyInt = 17
#adUnsignedSmallInt = 18
#adUnsignedInt = 19
#adUnsignedBigInt = 21
#adSingle = 4
#adDouble = 5
#adCurrency = 6
#adDecimal = 14
#adNumeric = 131
#adBoolean = 11
#adError = 10
#adUserDefined = 132
#adVariant = 12
#adIDispatch = 9
#adIUnknown = 13
#adGUID = 72
#adDate = 7
#adDBDate = 133
#adDBTime = 134
#adDBTimeStamp = 135
#adBSTR = 8
#adChar = 129
#adVarChar = 200
#adLongVarChar = 201
#adWChar = 130
#adVarWChar = 202
#adLongVarWChar = 203
#adBinary = 128
#adVarBinary = 204
#adLongVarBinary = 205
#adChapter = 136
#adFileTime = 64
#adPropVariant = 138
#adVarNumeric = 139
#adArray = $2000

;---- FieldAttributeEnum Values ----
#adFldMayDefer = $00000002
#adFldUpdatable = $00000004
#adFldUnknownUpdatable = $00000008
#adFldFixed = $00000010
#adFldIsNullable = $00000020
#adFldMayBeNull = $00000040
#adFldLong = $00000080
#adFldRowID = $00000100
#adFldRowVersion = $00000200
#adFldCacheDeferred = $00001000
#adFldIsChapter = $00002000
#adFldNegativeScale = $00004000
#adFldKeyColumn = $00008000
#adFldIsRowURL = $00010000
#adFldIsDefaultStream = $00020000
#adFldIsCollection = $00040000

;---- EditModeEnum Values ----
#adEditNone = $0000
#adEditInProgress = $0001
#adEditAdd = $0002
#adEditDelete = $0004

;---- RecordStatusEnum Values ----
#adRecOK = $0000000
#adRecNew = $0000001
#adRecModified = $0000002
#adRecDeleted = $0000004
#adRecUnmodified = $0000008
#adRecInvalid = $0000010
#adRecMultipleChanges = $0000040
#adRecPendingChanges = $0000080
#adRecCanceled = $0000100
#adRecCantRelease = $0000400
#adRecConcurrencyViolation = $0000800
#adRecIntegrityViolation = $0001000
#adRecMaxChangesExceeded = $0002000
#adRecObjectOpen = $0004000
#adRecOutOfMemory = $0008000
#adRecPermissionDenied = $0010000
#adRecSchemaViolation = $0020000
#adRecDBDeleted = $0040000

;---- GetRowsOptionEnum Values ----
#adGetRowsRest = -1

;---- PositionEnum Values ----
#adPosUnknown = -1
#adPosBOF = -2
#adPosEOF = -3

;---- BookmarkEnum Values ----
#adBookmarkCurrent = 0
#adBookmarkFirst = 1
#adBookmarkLast = 2

;---- MarshalOptionsEnum Values ----
#adMarshalAll = 0
#adMarshalModifiedOnly = 1

;---- AffectEnum Values ----
#adAffectCurrent = 1
#adAffectGroup = 2
#adAffectAllChapters = 4

;---- ResyncEnum Values ----
#adResyncUnderlyingValues = 1
#adResyncAllValues = 2

;---- CompareEnum Values ----
#adCompareLessThan = 0
#adCompareEqual = 1
#adCompareGreaterThan = 2
#adCompareNotEqual = 3
#adCompareNotComparable = 4

;---- FilterGroupEnum Values ----
#adFilterNone = 0
#adFilterPendingRecords = 1
#adFilterAffectedRecords = 2
#adFilterFetchedRecords = 3
#adFilterConflictingRecords = 5

;---- SearchDirectionEnum Values ----
#adSearchForward = 1
#adSearchBackward = -1

;---- PersistFormatEnum Values ----
#adPersistADTG = 0
#adPersistXML = 1

;---- StringFormatEnum Values ----
#adClipString = 2

;---- ConnectPromptEnum Values ----
#adPromptAlways = 1
#adPromptComplete = 2
#adPromptCompleteRequired = 3
#adPromptNever = 4

;---- ConnectModeEnum Values ----
#adModeUnknown = 0
#adModeRead = 1
#adModeWrite = 2
#adModeReadWrite = 3
#adModeShareDenyRead = 4
#adModeShareDenyWrite = 8
#adModeShareExclusive = $c
#adModeShareDenyNone = $10
#adModeRecursive = $400000

;---- RecordCreateOptionsEnum Values ----
#adCreateCollection = $00002000
#adCreateStructDoc = $80000000
#adCreateNonCollection = $00000000
#adOpenIfExists = $02000000
#adCreateOverwrite = $04000000
#adFailIfNotExists = -1

;---- RecordOpenOptionsEnum Values ----
#adOpenRecordUnspecified = -1
#adOpenOutput = $00800000
#adOpenAsync = $00001000
#adDelayFetchStream = $00004000
#adDelayFetchFields = $00008000
#adOpenExecuteCommand = $00010000

;---- IsolationLevelEnum Values ----
#adXactUnspecified = $ffffffff
#adXactChaos = $00000010
#adXactReadUncommitted = $00000100
#adXactBrowse = $00000100
#adXactCursorStability = $00001000
#adXactReadCommitted = $00001000
#adXactRepeatableRead = $00010000
#adXactSerializable = $00100000
#adXactIsolated = $00100000

;---- XactAttributeEnum Values ----
#adXactCommitRetaining = $00020000
#adXactAbortRetaining = $00040000

;---- PropertyAttributesEnum Values ----
#adPropNotSupported = $0000
#adPropRequired = $0001
#adPropOptional = $0002
#adPropRead = $0200
#adPropWrite = $0400

;---- ErrorValueEnum Values ----
#adErrProviderFailed = $bb8
#adErrInvalidArgument = $bb9
#adErrOpeningFile = $bba
#adErrReadFile = $bbb
#adErrWriteFile = $bbc
#adErrNoCurrentRecord = $bcd
#adErrIllegalOperation = $c93
#adErrCantChangeProvider = $c94
#adErrInTransaction = $cae
#adErrFeatureNotAvailable = $cb3
#adErrItemNotFound = $cc1
#adErrObjectInCollection = $d27
#adErrObjectNotSet = $d5c
#adErrDataConversion = $d5d
#adErrObjectClosed = $e78
#adErrObjectOpen = $e79
#adErrProviderNotFound = $e7a
#adErrBoundToCommand = $e7b
#adErrInvalidParamInfo = $e7c
#adErrInvalidConnection = $e7d
#adErrNotReentrant = $e7e
#adErrStillExecuting = $e7f
#adErrOperationCancelled = $e80
#adErrStillConnecting = $e81
#adErrInvalidTransaction = $e82
#adErrUnsafeOperation = $e84
#adwrnSecurityDialog = $e85
#adwrnSecurityDialogHeader = $e86
#adErrIntegrityViolation = $e87
#adErrPermissionDenied = $e88
#adErrDataOverflow = $e89
#adErrSchemaViolation = $e8a
#adErrSignMismatch = $e8b
#adErrCantConvertvalue = $e8c
#adErrCantCreate = $e8d
#adErrColumnNotOnThisRow = $e8e
#adErrURLIntegrViolSetColumns = $e8f
#adErrURLDoesNotExist = $e8f
#adErrTreePermissionDenied = $e90
#adErrInvalidURL = $e91
#adErrResourceLocked = $e92
#adErrResourceExists = $e93
#adErrCannotComplete = $e94
#adErrVolumeNotFound = $e95
#adErrOutOfSpace = $e96
#adErrResourceOutOfScope = $e97
#adErrUnavailable = $e98
#adErrURLNamedRowDoesNotExist = $e99
#adErrDelResOutOfScope = $e9a
#adErrPropInvalidColumn = $e9b
#adErrPropInvalidOption = $e9c
#adErrPropInvalidValue = $e9d
#adErrPropConflicting = $e9e
#adErrPropNotAllSettable = $e9f
#adErrPropNotSet = $ea0
#adErrPropNotSettable = $ea1
#adErrPropNotSupported = $ea2
#adErrCatalogNotSet = $ea3
#adErrCantChangeConnection = $ea4
#adErrFieldsUpdateFailed = $ea5
#adErrDenyNotSupported = $ea6
#adErrDenyTypeNotSupported = $ea7
#adErrProviderNotSpecified = $ea9
#adErrConnectionStringTooLong = $eaa

;---- ParameterAttributesEnum Values ----
#adParamSigned = $0010
#adParamNullable = $0040
#adParamLong = $0080

;---- ParameterDirectionEnum Values ----
#adParamUnknown = $0000
#adParamInput = $0001
#adParamOutput = $0002
#adParamInputOutput = $0003
#adParamReturnValue = $0004

;---- CommandTypeEnum Values ----
#adCmdUnknown = $0008
#adCmdText = $0001
#adCmdTable = $0002
#adCmdStoredProc = $0004
#adCmdFile = $0100
#adCmdTableDirect = $0200

;---- EventStatusEnum Values ----
#adStatusOK = $0000001
#adStatusErrorsOccurred = $0000002
#adStatusCantDeny = $0000003
#adStatusCancel = $0000004
#adStatusUnwantedEvent = $0000005

;---- EventReasonEnum Values ----
#adRsnAddNew = 1
#adRsnDelete = 2
#adRsnUpdate = 3
#adRsnUndoUpdate = 4
#adRsnUndoAddNew = 5
#adRsnUndoDelete = 6
#adRsnRequery = 7
#adRsnResynch = 8
#adRsnClose = 9
#adRsnMove = 10
#adRsnFirstChange = 11
#adRsnMoveFirst = 12
#adRsnMoveNext = 13
#adRsnMovePrevious = 14
#adRsnMoveLast = 15

;---- SchemaEnum Values ----
#adSchemaProviderSpecific = -1
#adSchemaAsserts = 0
#adSchemaCatalogs = 1
#adSchemaCharacterSets = 2
#adSchemaCollations = 3
#adSchemaColumns = 4
#adSchemaCheckConstraints = 5
#adSchemaConstraintColumnUsage = 6
#adSchemaConstraintTableUsage = 7
#adSchemaKeyColumnUsage = 8
#adSchemaReferentialConstraints = 9
#adSchemaTableConstraints = 10
#adSchemaColumnsDomainUsage = 11
#adSchemaIndexes = 12
#adSchemaColumnPrivileges = 13
#adSchemaTablePrivileges = 14
#adSchemaUsagePrivileges = 15
#adSchemaProcedures = 16
#adSchemaSchemata = 17
#adSchemaSQLLanguages = 18
#adSchemaStatistics = 19
#adSchemaTables = 20
#adSchemaTranslations = 21
#adSchemaProviderTypes = 22
#adSchemaViews = 23
#adSchemaViewColumnUsage = 24
#adSchemaViewTableUsage = 25
#adSchemaProcedureParameters = 26
#adSchemaForeignKeys = 27
#adSchemaPrimaryKeys = 28
#adSchemaProcedureColumns = 29
#adSchemaDBInfoKeywords = 30
#adSchemaDBInfoLiterals = 31
#adSchemaCubes = 32
#adSchemaDimensions = 33
#adSchemaHierarchies = 34
#adSchemaLevels = 35
#adSchemaMeasures = 36
#adSchemaProperties = 37
#adSchemaMembers = 38
#adSchemaTrustees = 39
#adSchemaFunctions = 40
#adSchemaActions = 41
#adSchemaCommands = 42
#adSchemaSets = 43

;---- FieldStatusEnum Values ----
#adFieldOK = 0
#adFieldCantConvertValue = 2
#adFieldIsNull = 3
#adFieldTruncated = 4
#adFieldSignMismatch = 5
#adFieldDataOverflow = 6
#adFieldCantCreate = 7
#adFieldUnavailable = 8
#adFieldPermissionDenied = 9
#adFieldIntegrityViolation = 10
#adFieldSchemaViolation = 11
#adFieldBadStatus = 12
#adFieldDefault = 13
#adFieldIgnore = 15
#adFieldDoesNotExist = 16
#adFieldInvalidURL = 17
#adFieldResourceLocked = 18
#adFieldResourceExists = 19
#adFieldCannotComplete = 20
#adFieldVolumeNotFound = 21
#adFieldOutOfSpace = 22
#adFieldCannotDeleteSource = 23
#adFieldReadOnly = 24
#adFieldResourceOutOfScope = 25
#adFieldAlreadyExists = 26
#adFieldPendingInsert = $10000
#adFieldPendingDelete = $20000
#adFieldPendingChange = $40000
#adFieldPendingUnknown = $80000
#adFieldPendingUnknownDelete = $100000

;---- SeekEnum Values ----
#adSeekFirstEQ = $1
#adSeekLastEQ = $2
#adSeekAfterEQ = $4
#adSeekAfter = $8
#adSeekBeforeEQ = $10
#adSeekBefore = $20

;---- ADCPROP_UPDATECRITERIA_ENUM Values ----
#adCriteriaKey = 0
#adCriteriaAllCols = 1
#adCriteriaUpdCols = 2
#adCriteriaTimeStamp = 3

;---- ADCPROP_ASYNCTHREADPRIORITY_ENUM Values ----
#adPriorityLowest = 1
#adPriorityBelowNormal = 2
#adPriorityNormal = 3
#adPriorityAboveNormal = 4
#adPriorityHighest = 5

;---- ADCPROP_AUTORECALC_ENUM Values ----
#adRecalcUpFront = 0
#adRecalcAlways = 1

;---- ADCPROP_UPDATERESYNC_ENUM Values ----
#adResyncNone = 0
#adResyncAutoIncrement = 1
#adResyncConflicts = 2
#adResyncUpdates = 4
#adResyncInserts = 8
#adResyncAll = 15

;---- MoveRecordOptionsEnum Values ----
#adMoveUnspecified = -1
#adMoveOverWrite = 1
#adMoveDontUpdateLinks = 2
#adMoveAllowEmulation = 4

;---- CopyRecordOptionsEnum Values ----
#adCopyUnspecified = -1
#adCopyOverWrite = 1
#adCopyAllowEmulation = 4
#adCopyNonRecursive = 2

;---- StreamTypeEnum Values ----
#adTypeBinary = 1
#adTypeText = 2

;---- LineSeparatorEnum Values ----
#adLF = 10
#adCR = 13
#adCRLF = -1

;---- StreamOpenOptionsEnum Values ----
#adOpenStreamUnspecified = -1
#adOpenStreamAsync = 1
#adOpenStreamFromRecord = 4

;---- StreamWriteEnum Values ----
#adWriteChar = 0
#adWriteLine = 1

;---- SaveOptionsEnum Values ----
#adSaveCreateNotExist = 1
#adSaveCreateOverWrite = 2

;---- FieldEnum Values ----
#adDefaultStream = -1
#adRecordURL = -2

;---- StreamReadEnum Values ----
#adReadAll = -1
#adReadLine = -2

;---- RecordTypeEnum Values ----
#adSimpleRecord = 0
#adCollectionRecord = 1
#adStructDoc = 2

; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 10
; Folding = -
; HideErrorLog