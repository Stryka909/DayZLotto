'#########################################################################
'#########    				DayZ Lottery							######
'#########	Developed by https://twitter.com/NullsecSlayer			######
'#########################################################################
'This script will insert random load out string in to 'instance' table resulting 
'in survivors spawning in with random inventory items. I wrote this script to support
' Bliss hive schema .30 and above.
'Instruction:
' 1) Modify DB connection parameters
' 2) Modify the loadout values if you want something different,
' 3) Use BEC scheduler or windows schedule to run this script
' Notes: To add more loadouts simply add new dLoadout.Add "[inventry]","[backpack]". Also don't forget to use double "" to escape " character in VBS. 
Option Explicit
Dim dLoadout
Set dLoadout = CreateObject("Scripting.Dictionary")
'--------------- EDIT BELOW ------------------
Const DBServer = "localhost" 'IP or Hostname of the DayZ MySQL server
Const DBName = "dayz" 'name of the hive DB
Const DBUser = "dayz" 'user name for the dayz DB
Const DBPass = "CHANGEME" 'DB password
Const DayZInstance = "1" 'instance of the dayz server

'If you screw up the load-oup string it will inject bad things in to your DB, don't cry like a bambi test your shit first.
dLoadout.Add "[[""AK_74""],[""30Rnd_545x39_AK"",""30Rnd_545x39_AK"",""30Rnd_545x39_AK""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""AKS_74_U""],[""30Rnd_545x39_AK"",""30Rnd_545x39_AK"",""30Rnd_545x39_AK""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""AKS_74_kobra""],[""30Rnd_545x39_AK"",""30Rnd_545x39_AK"",""30Rnd_545x39_AK""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""AK_47_M""],[""30Rnd_762x39_AK47"",""30Rnd_762x39_AK47"",""30Rnd_762x39_AK47""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""bizon_silenced""],[""64Rnd_9x19_SD_Bizon"",""64Rnd_9x19_SD_Bizon"",""64Rnd_9x19_SD_Bizon""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""Huntingrifle""],[""5x_22_LR_17_HMR"",""5x_22_LR_17_HMR"",""5x_22_LR_17_HMR"",""5x_22_LR_17_HMR""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""DMR""],[""20Rnd_762x51_DMR"",""20Rnd_762x51_DMR"",""20Rnd_762x51_DMR""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""FN_FAL""],[""20Rnd_762x51_FNFAL"",""20Rnd_762x51_FNFAL"",""20Rnd_762x51_FNFAL""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""FN_FAL_ANPVS4""],[""20Rnd_762x51_FNFAL"",""20Rnd_762x51_FNFAL"",""20Rnd_762x51_FNFAL""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""LeeEnfield""],[""10x_303"",""10x_303"",""10x_303"",""10x_303""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""M1014""],[""8Rnd_B_Beneli_74Slug"",""8Rnd_B_Beneli_74Slug"",""8Rnd_B_Beneli_74Slug""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""M14_EP1""],[""20Rnd_762x51_DMR"",""20Rnd_762x51_DMR"",""20Rnd_762x51_DMR"",""20Rnd_762x51_DMR""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""M16A2""],[""30Rnd_556x45_Stanag"",""30Rnd_556x45_Stanag"",""30Rnd_556x45_Stanag""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""M16A2GL""],[""30Rnd_556x45_Stanag"",""1Rnd_HE_M203"",""30Rnd_556x45_Stanag"",""1Rnd_HE_M203"",""30Rnd_556x45_Stanag"",""1Rnd_HE_M203""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""M16A4_ACG""],[""30Rnd_556x45_Stanag"",""30Rnd_556x45_Stanag"",""30Rnd_556x45_Stanag""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""M24""],[""5Rnd_762x51_M24"",""5Rnd_762x51_M24"",""5Rnd_762x51_M24"",""5Rnd_762x51_M24""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""M240""],[""100Rnd_762x51_M240"",""100Rnd_762x51_M240""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""M249""],[""200Rnd_556x45_M249"",""200Rnd_556x45_M249""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""M4A1""],[""30Rnd_556x45_Stanag"",""30Rnd_556x45_Stanag"",""30Rnd_556x45_Stanag""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""M4A1_HWS_GL""],[""30Rnd_556x45_Stanag"",""1Rnd_HE_M203"",""30Rnd_556x45_Stanag"",""1Rnd_HE_M203"",""30Rnd_556x45_Stanag"",""1Rnd_HE_M203""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""M4A1_Aim""],[""30Rnd_556x45_Stanag"",""30Rnd_556x45_Stanag"",""30Rnd_556x45_Stanag""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""M4A1_AIM_SD_camo""],[""30Rnd_556x45_StanagSD"",""30Rnd_556x45_StanagSD"",""30Rnd_556x45_StanagSD""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""M4A3_CCO_EP1""],[""30Rnd_556x45_Stanag"",""30Rnd_556x45_Stanag"",""30Rnd_556x45_Stanag""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""Mk_48""],[""100Rnd_762x51_M240"",""100Rnd_762x51_M240""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""MP5A5""],[""30rnd_9x19_MP5"",""30rnd_9x19_MP5"",""30rnd_9x19_MP5""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""MP5SD""],[""30rnd_9x19_MP5SD"",""30rnd_9x19_MP5SD"",""30rnd_9x19_MP5SD""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""Remington870_lamp""],[""8Rnd_B_Beneli_74Slug"",""8Rnd_B_Beneli_74Slug"",""8Rnd_B_Beneli_74Slug""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""SVD_CAMO""],[""10Rnd_762x54_SVD"",""10Rnd_762x54_SVD"",""10Rnd_762x54_SVD""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""Winchester1866""],[""8Rnd_B_Beneli_74Slug"",""8Rnd_B_Beneli_74Slug"",""8Rnd_B_Beneli_74Slug""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""glock17_EP1""],[""17Rnd_9x19_glock17"",""17Rnd_9x19_glock17"",""17Rnd_9x19_glock17"",""17Rnd_9x19_glock17""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""Colt1911""],[""7Rnd_45ACP_1911"",""7Rnd_45ACP_1911"",""7Rnd_45ACP_1911"",""7Rnd_45ACP_1911""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""M9""],[""15Rnd_9x19_M9"",""15Rnd_9x19_M9"",""15Rnd_9x19_M9"",""15Rnd_9x19_M9""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""M9SD""],[""15Rnd_9x19_M9SD"",""15Rnd_9x19_M9SD"",""15Rnd_9x19_M9SD"",""15Rnd_9x19_M9SD""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""Makarov""],[""8Rnd_9x18_Makarov"",""8Rnd_9x18_Makarov"",""8Rnd_9x18_Makarov"",""8Rnd_9x18_Makarov""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""revolver_EP1""],[""6Rnd_45ACP"",""6Rnd_45ACP"",""6Rnd_45ACP"",""6Rnd_45ACP""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[[""UZI_EP1""],[""30Rnd_9x19_UZI"",""30Rnd_9x19_UZI"",""30Rnd_9x19_UZI"",""30Rnd_9x19_UZI""]]", "[""DZ_Patrol_Pack_EP1"",[[],[]],[[],[]]]"
dLoadout.Add "[]", "[]" 'Some really unlucky SOB will get nothing. World is not fair, deals with it!
'--------------- STOP EDIT -------------------

Dim oDBCon, arrInventry, arrBackpack, intPowerball

arrInventry = dLoadout.Keys
arrBackpack = dLoadout.Items
intPowerball= RndInt(LBound(arrInventry), UBound(arrInventry))
Set oDBCon = CreateObject("ADODB.Connection")
	oDBCon.Open "Driver={MySQL ODBC 5.2w Driver};Server=" & DBServer & ";Database=" & DBName & ";User=" & DBUser & ";Password=" & DBPass & ";Option=3;"
	oDBCon.Execute "UPDATE instance SET inventory='" & arrInventry(intPowerball) & "', backpack='" & arrBackpack(intPowerball) & "' WHERE id=" & DayZInstance
	'WScript.Echo "UPDATE instance SET inventory='" & arrInventry(intPowerball) & "', backpack='" & arrBackpack(intPowerball) & "' WHERE id=" & DayZInstance
	oDBCon.Close
Set oDBCon = Nothing
Set dLoadout = Nothing


Function RndInt(ByVal myMin, ByVal myMax)
Randomize Timer
RndInt = Int((myMax - myMin + 1)*Rnd() + myMin)
End Function






'---- CursorTypeEnum Values ----
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

'---- CursorOptionEnum Values ----
Const adHoldRecords = &H00000100
Const adMovePrevious = &H00000200
Const adAddNew = &H01000400
Const adDelete = &H01000800
Const adUpdate = &H01008000
Const adBookmark = &H00002000
Const adApproxPosition = &H00004000
Const adUpdateBatch = &H00010000
Const adResync = &H00020000
Const adNotify = &H00040000

'---- LockTypeEnum Values ----
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

'---- ExecuteOptionEnum Values ----
Const adRunAsync = &H00000010

'---- ObjectStateEnum Values ----
Const adStateClosed = &H00000000
Const adStateOpen = &H00000001
Const adStateConnecting = &H00000002
Const adStateExecuting = &H00000004

'---- CursorLocationEnum Values ----
Const adUseServer = 2
Const adUseClient = 3


'---- DataTypeEnum Values ----
Const adEmpty = 0
Const adTinyInt = 16
Const adSmallInt = 2
Const adInteger = 3
Const adBigInt = 20
Const adUnsignedTinyInt = 17
Const adUnsignedSmallInt = 18
Const adUnsignedInt = 19
Const adUnsignedBigInt = 21
Const adSingle = 4
Const adDouble = 5
Const adCurrency = 6
Const adDecimal = 14
Const adNumeric = 131
Const adBoolean = 11
Const adError = 10
Const adUserDefined = 132
Const adVariant = 12
Const adIDispatch = 9
Const adIUnknown = 13
Const adGUID = 72
Const adDate = 7
Const adDBDate = 133
Const adDBTime = 134
Const adDBTimeStamp = 135
Const adBSTR = 8
Const adChar = 129
Const adVarChar = 200
Const adLongVarChar = 201
Const adWChar = 130
Const adVarWChar = 202
Const adLongVarWChar = 203
Const adBinary = 128
Const adVarBinary = 204
Const adLongVarBinary = 205

'---- FieldAttributeEnum Values ----
Const adFldMayDefer = &H00000002
Const adFldUpdatable = &H00000004
Const adFldUnknownUpdatable = &H00000008
Const adFldFixed = &H00000010
Const adFldIsNullable = &H00000020
Const adFldMayBeNull = &H00000040
Const adFldLong = &H00000080
Const adFldRowID = &H00000100
Const adFldRowVersion = &H00000200
Const adFldCacheDeferred = &H00001000

'---- EditModeEnum Values ----
Const adEditNone = &H0000
Const adEditInProgress = &H0001
Const adEditAdd = &H0002
Const adEditDelete = &H0004

'---- RecordStatusEnum Values ----
Const adRecOK = &H0000000
Const adRecNew = &H0000001
Const adRecModified = &H0000002
Const adRecDeleted = &H0000004
Const adRecUnmodified = &H0000008
Const adRecInvalid = &H0000010
Const adRecMultipleChanges = &H0000040
Const adRecPendingChanges = &H0000080
Const adRecCanceled = &H0000100
Const adRecCantRelease = &H0000400
Const adRecConcurrencyViolation = &H0000800
Const adRecIntegrityViolation = &H0001000
Const adRecMaxChangesExceeded = &H0002000
Const adRecObjectOpen = &H0004000
Const adRecOutOfMemory = &H0008000
Const adRecPermissionDenied = &H0010000
Const adRecSchemaViolation = &H0020000
Const adRecDBDeleted = &H0040000

'---- GetRowsOptionEnum Values ----
Const adGetRowsRest = -1

'---- PositionEnum Values ----
Const adPosUnknown = -1
Const adPosBOF = -2
Const adPosEOF = -3

'---- enum Values ----
Const adBookmarkCurrent = 0
Const adBookmarkFirst = 1
Const adBookmarkLast = 2

'---- MarshalOptionsEnum Values ----
Const adMarshalAll = 0
Const adMarshalModifiedOnly = 1

'---- AffectEnum Values ----
Const adAffectCurrent = 1
Const adAffectGroup = 2
Const adAffectAll = 3

'---- FilterGroupEnum Values ----
Const adFilterNone = 0
Const adFilterPendingRecords = 1
Const adFilterAffectedRecords = 2
Const adFilterFetchedRecords = 3
Const adFilterPredicate = 4

'---- SearchDirection Values ----
Const adSearchForward = 1
Const adSearchBackward = -1

'---- ConnectPromptEnum Values ----
Const adPromptAlways = 1
Const adPromptComplete = 2
Const adPromptCompleteRequired = 3
Const adPromptNever = 4

'---- ConnectModeEnum Values ----
Const adModeUnknown = 0
Const adModeRead = 1
Const adModeWrite = 2
Const adModeReadWrite = 3
Const adModeShareDenyRead = 4
Const adModeShareDenyWrite = 8
Const adModeShareExclusive = &Hc
Const adModeShareDenyNone = &H10

'---- IsolationLevelEnum Values ----
Const adXactUnspecified = &Hffffffff
Const adXactChaos = &H00000010
Const adXactReadUncommitted = &H00000100
Const adXactBrowse = &H00000100
Const adXactCursorStability = &H00001000
Const adXactReadCommitted = &H00001000
Const adXactRepeatableRead = &H00010000
Const adXactSerializable = &H00100000
Const adXactIsolated = &H00100000

'---- XactAttributeEnum Values ----
Const adXactCommitRetaining = &H00020000
Const adXactAbortRetaining = &H00040000

'---- PropertyAttributesEnum Values ----
Const adPropNotSupported = &H0000
Const adPropRequired = &H0001
Const adPropOptional = &H0002
Const adPropRead = &H0200
Const adPropWrite = &H0400

'---- ErrorValueEnum Values ----
Const adErrInvalidArgument = &Hbb9
Const adErrNoCurrentRecord = &Hbcd
Const adErrIllegalOperation = &Hc93
Const adErrInTransaction = &Hcae
Const adErrFeatureNotAvailable = &Hcb3
Const adErrItemNotFound = &Hcc1
Const adErrObjectInCollection = &Hd27
Const adErrObjectNotSet = &Hd5c
Const adErrDataConversion = &Hd5d
Const adErrObjectClosed = &He78
Const adErrObjectOpen = &He79
Const adErrProviderNotFound = &He7a
Const adErrBoundToCommand = &He7b
Const adErrInvalidParamInfo = &He7c
Const adErrInvalidConnection = &He7d
Const adErrStillExecuting = &He7f
Const adErrStillConnecting = &He81

'---- ParameterAttributesEnum Values ----
Const adParamSigned = &H0010
Const adParamNullable = &H0040
Const adParamLong = &H0080

'---- ParameterDirectionEnum Values ----
Const adParamUnknown = &H0000
Const adParamInput = &H0001
Const adParamOutput = &H0002
Const adParamInputOutput = &H0003
Const adParamReturnValue = &H0004

'---- CommandTypeEnum Values ----
Const adCmdUnknown = &H0008
Const adCmdText = &H0001
Const adCmdTable = &H0002
Const adCmdStoredProc = &H0004

'---- SchemaEnum Values ----
Const adSchemaProviderSpecific = -1
Const adSchemaAsserts = 0
Const adSchemaCatalogs = 1
Const adSchemaCharacterSets = 2
Const adSchemaCollations = 3
Const adSchemaColumns = 4
Const adSchemaCheckConstraints = 5
Const adSchemaConstraintColumnUsage = 6
Const adSchemaConstraintTableUsage = 7
Const adSchemaKeyColumnUsage = 8
Const adSchemaReferentialContraints = 9
Const adSchemaTableConstraints = 10
Const adSchemaColumnsDomainUsage = 11
Const adSchemaIndexes = 12
Const adSchemaColumnPrivileges = 13
Const adSchemaTablePrivileges = 14
Const adSchemaUsagePrivileges = 15
Const adSchemaProcedures = 16
Const adSchemaSchemata = 17
Const adSchemaSQLLanguages = 18
Const adSchemaStatistics = 19
Const adSchemaTables = 20
Const adSchemaTranslations = 21
Const adSchemaProviderTypes = 22
Const adSchemaViews = 23
Const adSchemaViewColumnUsage = 24
Const adSchemaViewTableUsage = 25
Const adSchemaProcedureParameters = 26
Const adSchemaForeignKeys = 27
Const adSchemaPrimaryKeys = 28
Const adSchemaProcedureColumns = 29

'---- FSO Constants
Const ForAppending = 8
Const ForReading = 1
Const ForWriting = 2