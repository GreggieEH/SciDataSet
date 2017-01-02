// dispids.h
// dispatch ids

#pragma once

// properties
#define			DISPID_DataProgID				0x0100
#define			DISPID_RadianceAvailable		0x0101
#define			DISPID_IrradianceAvailable		0x0102
#define			DISPID_AnalysisMode				0x0103
#define			DISPID_FileName					0x0104
#define			DISPID_AmCalibration			0x0105
#define			DISPID_RadianceWaves			0x0106
#define			DISPID_Radiance					0x0107
#define			DISPID_Irradiance				0x0108
#define			DISPID_HaveOPT					0x0109
#define			DISPID_OPTWaves					0x010a
#define			DISPID_OPTValues				0x010b
#define			DISPID_ScanType					0x010c
#define			DISPID_Comment					0x010d
#define			DISPID_DefaultUnits				0x010e
#define			DISPID_CountsAvailable			0x010f
#define			DISPID_HaveCalibrationStandard	0x0110
#define			DISPID_DefaultTitle				0x0111
#define			DISPID_ArrayScanning			0x0112
#define			DISPID_NumPeaks					0x0113
#define			DISPID_PeakPosition				0x0114
#define			DISPID_PeakHeight				0x0115
#define			DISPID_PeakWidth				0x0116

// methods
#define			DISPID_ReadDataFile				0x0120
#define			DISPID_WriteDataFile			0x0121
#define			DISPID_GetWavelengths			0x0122
#define			DISPID_GetSpectra				0x0123
#define			DISPID_AddValue					0x0124
#define			DISPID_ViewSetup				0x0125
#define			DISPID_GetArraySize				0x0126
#define			DISPID_SetCurrentExp			0x0127
#define			DISPID_calcRadiance				0x0128
#define			DISPID_setObject				0x0129
#define			DISPID_GetCreateTime			0x012b
#define			DISPID_FormOutFile				0x012c
#define			DISPID_CanSetAnalysisMode		0x012d
#define			DISPID_GetDataValue				0x012e
#define			DISPID_SetDataValue				0x012f
#define			DISPID_LoadFromString			0x0130
#define			DISPID_FormOpticalTransferFile	0x0131
//#define			DISPID_SetCalibrationWavelengths	0x0132
#define			DISPID_SetCalibrationData		0x0133
#define			DISPID_FileNameFromPath			0x0134
#define			DISPID_NewObjectSetup			0x0135
#define			DISPID_SelectCalibrationStandard	0x0136
#define			DISPID_ExtraScanData			0x0137
#define			DISPID_GetGratingScan			0x0138
#define			DISPID_ExportToCSV				0x0139
#define			DISPID_GetADGainFactor			0x013a
#define			DISPID_SetArrayData				0x013b
#define			DISPID_InitTempData				0x013c
#define			DISPID_FindPeaks				0x013d
#define			DISPID_DisplayPeaks				0x013e
#define			DISPID_IntegrateSpectra			0x013f
#define			DISPID_SmoothSpectra			0x1120
#define			DISPID_CalcAverage				0x1121

// events
#define			DISPID_requestDispersion		0x0140
#define			DISPID_requestAnalysisModeIndex	0x0141
#define			DISPID_requestAnalysisModeString	0x0142
#define			DISPID_requestOpticalTransfer	0x0143
#define			DISPID_RequestOPTArraySize		0x0144
#define			DISPID_RequestOPTDataValue		0x0145
#define			DISPID_RequestSetupDisplay		0x0146
#define			DISPID_updateOptFile			0x0147
#define			DISPID_requestADInfoString		0x0148
#define			DISPID_requestCalibrationScan	0x0149
#define			DISPID_requestCalibrationGain	0x014a
#define			DISPID_requestWorkingDirectory	0x014b
#define			DISPID_browseForWorkingDirectory	0x014c
#define			DISPID_setWorkingDirectory		0x014d