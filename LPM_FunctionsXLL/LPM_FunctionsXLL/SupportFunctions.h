///***************************************************************************
// File:				SupportFunctions.H
//
// Purpose:			Header file for TracerLPMfunctions.c
// 
// Platform:    Microsoft Windows
//
// Development Team: Bryant Jurgens
//
// Updated by Microsoft Product Support Services, Windows Developer Support.
// From the Microsoft Excel Developer's Kit, Version 5
// Copyright (c) 1996 Microsoft Corporation. All rights reserved.
///***************************************************************************

// 
// Function prototypes
//
void cwCenter(HWND, int);
BOOL CALLBACK DIALOGMsgProc(HWND hWndDlg, UINT message, WPARAM wParam, LPARAM lParam);
BOOL GetHwnd(HWND * pHwnd);
int lpwstricmp(LPWSTR s, LPWSTR t);
double DispInt(double X, double Tau, double Lambda, double DP);
double DispInt2(double X, double Tau, double Lambda, double DP);
double DispInt3(double X, double Tau, double Lambda, double DP);
double gt_PFMint(double MinAge, double MaxAge, double Tau, double UZtime);
double gt_EMMint(double MinAge, double MaxAge, double Tau, double UZtime);
double gt_EPMint(double MinAge, double MaxAge, double Tau, double EPMratio, double UZtime);
double gt_DMint(double MinAge, double MaxAge, double Tau, double Lambda, double DP);
double gt_DMint2(double MinAge, double MaxAge, double Tau, double Lambda, double DP);
double MinTimeInc(FP DateRange[]);
int sign(const double X);
double gt_FDMint(double MinAge, double& MaxAge, double& Alpha, double& Tau, double& DP, double& Lambda, double& UZtime);
double gtFDMaux(double a, double b, double epsilon, double S, double fa, double fb, double fc, int bottom, double Alpha, double Tau, double DP);
double ReturnExcelGamma(double age, double alpha, double beta);
double gt_GAMint(double MinAge, double MaxAge, double Tau, double Alpha, double UZtime);
double adaptiveSimpsons(double& T, double& Alpha, double& Tau, double& DP, double& Lambda);
double adaptiveSimpsonsAux(double a, double b, double epsilon, double S, double fa, double fb, double fc, int bottom, double& X, double& Alpha);
double ReturnLambdaCorrection(double MinAge, double MaxAge, double Lambda);
//
// identifier for controls
//
#define FREE_SPACE                  104
#define EDIT                        101
#define TEST_EDIT                   106