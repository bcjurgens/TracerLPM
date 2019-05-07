// TracerLPMfunctions.c : Defines the exported functions for the DLL application.
//
//
#include <windows.h>
#include <xlcall.h>
#include <framewrk.h>
//#include <math.h>
#include <complex>
#include <cmath>
//#pragma inline_recursion(on)
using namespace std; 

//#include "SupportFunctions.h"

// Global Variables
//
HWND g_hWndMain = NULL;
HANDLE g_hInst = NULL;
XCHAR g_szBuffer[20] = L"";
//
// Syntax of the Register Command:
//      REGISTER(module_text, procedure, type_text, function_text, 
//               argument_text, macro_type, category, shortcut_text,
//               help_topic, function_help, argument_help1, argument_help2,...)
//
//
// g_rgWorksheetFuncs will use only the first 11 arguments of 
// the Register function.
//
// This is a table of all the worksheet functions exported by this module.
// These functions are all registered (in xlAutoOpen) when you
// open the XLL. Before every string, leave a space for the
// byte count. The format of this table is the same as 
// arguments two through eleven of the REGISTER function.
// g_rgWorksheetFuncsRows define the number of rows in the table. The
// g_rgWorksheetFuncsCols represents the number of columns in the table.
#define g_rgWorksheetFuncsRows 43 //changed from 44 to match removal of C14 functions
#define g_rgWorksheetFuncsCols 24
static LPWSTR g_rgWorksheetFuncs
[g_rgWorksheetFuncsRows][g_rgWorksheetFuncsCols] =
{
	{ L"EMM",
		L"BKKBBBBMMM$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"EMM",
		L"Date range,Tracer input range,Tau, Sample date, Lambda, UZtt, Is3He(trit), Is3H, Is3H/3Ho",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet tracer concentration from the exponential mixing model",   
		L"The range corresponding to the tracer input dates (in decimal years).",                   
		L"The range corresponding to the tracer input concentration data.",
		L"The mean age in years - number",
		L"The sample date in decimal years - number",
		L"The decay rate of the tracer [ln(2)/half-life] - number",
		L"The unsaturated zone travel time in years - number",
		L"Is the tracer tritiogenic helium-3 - True or False",
		L"Is the tracer initial tritium - True or False",
		L"Is the tracer tritium/initial tritium - True or False"
	},
	{ L"GAM",
		L"BKKBBBBBAAA",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"GAM",
		L"Date range,Tracer input range,Tau, Sample date(decimal year), Decay Rate (lambda), Alpha, Unsat. zone travel time, Is helium-3, Is initial tritium, Is tritium/initial tritium ratio",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet tracer concentration from the exponential piston-flow model",   
		L"The range corresponding to the tracer input dates (in decimal years).",                   
		L"The range corresponding to the tracer input concentration data.",
		L"The mean age in years - number",
		L"The sample date in decimal years - number",
		L"The decay rate of the tracer [ln(2)/half-life] - number",
		L"The shape parameter of the GAMMA function - number",
		L"The unsaturated zone travel time in years - number",
		L"Is the tracer helium-3 - True or False",
		L"Is the tracer initial tritium - True or False",
		L"Is the tracer tritium/initial tritium - True or False"
	},
	{ L"FDM",
		L"BKKBBBBBBMMM$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"FDM",
		L"Date range,Tracer input range,Tau, Sample date, Lambda, Alpha, DP, UZtt, Is3He(trit), Is3H, Is3H/3Ho",
		//111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet tracer concentration from the fractional-dispersion model",   
		L"The range corresponding to the tracer input dates (in decimal years).",                   
		L"The range corresponding to the tracer input concentration data.",
		L"The mean age in years - number",
		L"The sample date in decimal years - number",
		L"The decay rate of the tracer [ln(2)/half-life] - number",
		L"The fractional-order derivative of the spatial dispersion term - number",
		L"The ratio of dispersion to advection or inverse of the Peclet number (D/vx) - number",
		L"The unsaturated zone travel time in years - number",
		L"Is the tracer tritiogenic helium-3 - True or False",
		L"Is the tracer initial tritium - True or False",
		L"Is the tracer tritium/initial tritium - True or False"
	},
	{ L"gt_FDM",
		L"BBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"gt_FDM",
		L"MinAge, MaxAge, Alpha, Tau, DP, UZtime",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the fraction of recharge for a specified age interval",   
		L"",                   
		L"MaxAge",
		L"The fractional-order derivative of the spatial dispersion term - number",
		L"The mean age in years - number",
		L"The dispersion parameter",
		L"Unsaturated zone travel time"
	},
	{ L"AlphaDensity",
		L"BBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"AlphaDensity",
		L"T, X, Alpha, Beta, C, Mu",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the alpha density at a time (T) and location (X)",   
		L"Time.",                   
		L"X",
		L"Alpha",
		L"Beta",
		L"C",
		L"Mu"
	},
	{ L"AlphaStablePDF",
		L"BBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"AlphaStablePDF",
		L"X, Alpha, Beta, C, Mu, a, b, n, Tol",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the something to do with alpha density",   
		L"Alpha",                   
		L"Beta",
		L"C",
		L"Mu",
		L"a",
		L"b",
		L"n",
		L"Tol"
	},
	{ L"PEM",
		L"BKKBBBBBBMMM$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"PEM",
		L"Date range,Tracer input range,Tau, Sample date, Lambda, PEM Upper ratio, PEM Lower ratio, UZtt, Is3He(trit), Is3H, Is3H/3Ho",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet tracer concentration from the exponential mixing model",   
		L"The range corresponding to the tracer input dates (in decimal years).",                   
		L"The range corresponding to the tracer input concentration data.",
		L"The mean age in years - number",
		L"The sample date in decimal years - number",
		L"The decay rate of the tracer [ln(2)/half-life] - number",
		L"The ratio of the distance between the water table and top of the screen and the distance between the top of the screen and the bottom of the aquifer - number",
		L"The ratio of the distance between the water table and bottom of the screen and the distance between the bottom of the screen and the bottom of the aquifer - number",
		L"The unsaturated zone travel time in years - number",
		L"Is the tracer tritiogenic helium-3 - True or False",
		L"Is the tracer initial tritium - True or False",
		L"Is the tracer tritium/initial tritium - True or False"
	},
	{ L"PFM",
		L"BKKBBBBMMM$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"PFM",
		L"Date range,Tracer input range,Tau, Sample date, Lambda, UZtt, Is3He(trit), Is3H, Is3H/3Ho",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet tracer concentration from the piston-flow model",   
		L"The range corresponding to the tracer input dates (in decimal years).",                   
		L"The range corresponding to the tracer input concentration data.",
		L"The mean age in years - number",
		L"The sample date in decimal years - number",
		L"The decay rate of the tracer [ln(2)/half-life] - number",
		L"The unsaturated zone travel time in years - number",
		L"Is the tracer tritiogenic helium-3 - True or False",
		L"Is the tracer initial tritium - True or False",
		L"Is the tracer tritium/initial tritium - True or False"                   
	},
	{ L"EPM",
		L"BKKBBBBBMMM$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"EPM",
		L"Date range,Tracer input range,Tau, Sample date, Lambda, EPM ratio, UZtt, Is3He(trit), Is3H, Is3H/3Ho",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet tracer concentration from the exponential piston-flow model",   
		L"The range corresponding to the tracer input dates (in decimal years).",                   
		L"The range corresponding to the tracer input concentration data.",
		L"The mean age in years - number",
		L"The sample date in decimal years - number",
		L"The decay rate of the tracer [ln(2)/half-life] - number",
		L"The ratio of the recharge length to the length of piston-flow travel - number",
		L"The unsaturated zone travel time in years - number",
		L"Is the tracer tritiogenic helium-3 - True or False",
		L"Is the tracer initial tritium - True or False",
		L"Is the tracer tritium/initial tritium - True or False"
	},
	{ L"gt_DM",
		L"BBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"gt_DM",
		L"Min Age, Max Age, Tau, DP, Unsat. zone travel time",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the fractional contribution of flow for a given age interval using the dispersion model",   
		L"The minimum age of the age interval - number",  
		L"The maximum age of the age interval - number",   
		L"The mean age in years - number",
		L"The ratio of dispersion to advection or inverse of the Peclet number (D/vx) - number",
		L"The unsaturated zone travel time in years - number"   
	},
	{ L"DM",
		L"BKKBBBBBMMM$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"DM",
		L"Date range,Tracer input range,Tau, Sample date, Lambda, DP, UZtt, Is3He(trit), Is3H, Is3H/3Ho",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet tracer concentration from the dispersion model",   
		L"The range corresponding to the tracer input dates (in decimal years).",                   
		L"The range corresponding to the tracer input concentration data.",
		L"The mean age in years - number",
		L"The sample date in decimal years - number",
		L"The decay rate of the tracer [ln(2)/half-life] - number",
		L"The ratio of dispersion to advection or inverse of the Peclet number (D/vx) - number",
		L"The unsaturated zone travel time in years - number",
		L"Is the tracer tritiogenic helium-3 - True or False",
		L"Is the tracer initial tritium - True or False",
		L"Is the tracer tritium/initial tritium - True or False"
	},
	{ L"PFM_He4",
		L"BBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"PFM_He4",
		L"Uranium, Thorium, Phi, Rho, Tau, HeSolRate",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet Helium-4 concentration from the piston-flow model",   
		L"Uranium concentration of aquifer in parts per million - number",                   
		L"Thorium concentration of aquifer in parts per million - number",
		L"Porosity - number",
		L"Bulk density - number",
		L"The mean age in years - number",              
		L"Helium solution rate (cc@STP/gramH2O/year) - number"
	},
	{ L"EMM_He4",
		L"BBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"EMM_He4",
		L"Uranium, Thorium, Phi, Rho, Tau, HeSolRate",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet Helium-4 concentration from the exponential mixing model",   
		L"Uranium concentration of aquifer in parts per million - number",                   
		L"Thorium concentration of aquifer in parts per million - number",
		L"Porosity - number",
		L"Bulk density - number",
		L"The mean age in years - number",
		L"Helium solution rate (cc@STP/gramH2O/year) - number"
	},
	{ L"PEM_He4",
		L"BBBBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"PEM_He4",
		L"Uranium, Thorium, Phi, Rho, Tau, PEM Uratio, PEM Lratio, HeSolRate",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet Helium-4 concentration from the partial exponential mixing model",   
		L"Uranium concentration of aquifer in parts per million - number",                   
		L"Thorium concentration of aquifer in parts per million - number",
		L"Porosity - number",
		L"Bulk density - number",
		L"The mean age in years - number",
		L"The ratio of the distance between the water table and top of the screen and the distance between the top of the screen and the bottom of the aquifer - number",
		L"The ratio of the distance between the water table and bottom of the screen and the distance between the bottom of the screen and the bottom of the aquifer - number",
		L"Helium solution rate (cc@STP/gramH2O/year) - number"
	},
	{ L"GAM_He4",
		L"BBBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"GAM_He4",
		L"Uranium (ppm), Thorium (ppm), Porosity (phi), Sediment density (rho), Tau, Alpha, Helium solution rate",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet Helium-4 concentration from the exponential mixing model",   
		L"Uranium concentration of aquifer in parts per million - number",                   
		L"Thorium concentration of aquifer in parts per million - number",
		L"Porosity - number",
		L"Sediment density - number",
		L"The mean age in years - number",
		L"The shape parameter for the GAMMA function - number",
		L"Helium solution rate (per year) - number"
	},
	{ L"EPM_He4",
		L"BBBBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"EPM_He4",
		L"Uranium, Thorium, Phi, Rho, Tau, EPM ratio, HeSolRate",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet Helium-4 concentration from the exponentialpiston-flow model",   
		L"Uranium concentration of aquifer in parts per million - number",                   
		L"Thorium concentration of aquifer in parts per million - number",
		L"Porosity - number",
		L"Bulk density - number",
		L"The mean age in years - number",
		L"The ratio of the recharge length to the length of piston-flow travel - number",
		L"Helium solution rate (cc@STP/gramH2O/year) - number"
	},
	{ L"DM_He4",
		L"BBBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"DM_He4",
		L"Uranium, Thorium, Phi, Rho, Tau, DP, HeSolRate",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet Helium-4 concentration from the dispersion model",   
		L"Uranium concentration of aquifer in parts per million - number",                   
		L"Thorium concentration of aquifer in parts per million - number",
		L"Porosity - number",
		L"Bulk density - number",
		L"The mean age in years - number",
		L"The ratio of dispersion to advection or inverse of the Peclet number (D/vx) - number",
		L"Helium solution rate (cc@STP/gramH2O/year) - number"
	},
	{ L"gt_PFM",
		L"BBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"gt_PFM",
		L"Min Age, Max Age, Tau, Unsat. zone travel time",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the fractional contribution of flow for a given age interval using the piston-flow model",   
		L"The minimum age of the age interval - number",  
		L"The maximum age of the age interval - number",   
		L"The mean age in years - number",
		L"The unsaturated zone travel time in years - number"
	},
	{ L"gt_EMM",
		L"BBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"gt_EMM",
		L"Min Age, Max Age, Tau, Unsat. zone travel time",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the fractional contribution of flow for a given age interval using the exponential mixing model",   
		L"The minimum age of the age interval - number",  
		L"The maximum age of the age interval - number",   
		L"The mean age in years - number",
		L"The unsaturated zone travel time in years - number"                  
	},
	{ L"gt_PEM",
		L"BBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"gt_PEM",
		L"Min Age, Max Age, Tau, PEM Upper ratio, PEM Lower ratio, Unsat. zone travel time",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the fractional contribution of flow for a given age interval using the partial exponential mixing model",   
		L"The minimum age of the age interval - number",  
		L"The maximum age of the age interval - number",   
		L"The mean age in years - number",
		L"The ratio of the distance between the water table and top of the screen and the distance between the top of the screen and the bottom of the aquifer - number",
		L"The ratio of the distance between the water table and bottom of the screen and the distance between the bottom of the screen and the bottom of the aquifer - number",
		L"The unsaturated zone travel time in years - number"                  
	},
	{ L"gt_EPM",
		L"BBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"gt_EPM",
		L"Min Age, Max Age, Tau, EPM ratio, UZtt",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the fractional contribution of flow for a given age interval using the exponential piston-flow model",   
		L"The minimum age of the age interval - number",  
		L"The maximum age of the age interval - number",   
		L"The mean age in years - number",
		L"The ratio of the recharge length to the length of piston-flow travel - number",
		L"The unsaturated zone travel time in years - number"
	},
	{L"gt_GAM",
		L"BBBBBB",
		L"gt_GAM",
		L"Min Age, Max Age, Tau, Alpha, UZtime",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the fractional contribution of flow for a given age interval using the GAM model",   
		L"The minimum age of the age interval - number",  
		L"The maximum age of the age interval - number",   
		L"The mean age in years - number",
		L"The shape parameter of the GAMMA function - number",
		L"The unsaturated zone travel time in years - number"
	},
	{ L"BMM_DM",
		L"BKKBBBBBBBMMMBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"BMM_DM",
		L"Date range, Tracer input range, Tau, Sample date, Lambda, DP, UZtt, MixFrac, Conc2ndComp, Is3He(trit), Is3H, Is3H/3Ho, DIC1, DIC2",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet tracer concentration from a binary mixture with a dispersion model for the 1st component",
		L"The range corresponding to the tracer input dates (in decimal years).",    
		L"The range corresponding to the tracer input concentration data.",
		L"The mean age of the 1st component in years - number",
		L"The sample date in decimal years - number",
		L"The decay rate of the tracer [ln(2)/half-life] - number",
		L"The ratio of dispersion to advection or inverse of the Peclet number (D/vx) - number",
		L"The unsaturated zone travel time in years - number",
		L"Mixing fraction of the 1st component - number between 0 and 1",
		L"Conc2ndComp, or formula for 2nd component - number",
		L"Is the tracer tritiogenic helium-3 - True or False",
		L"Is the tracer initial tritium - True or False",
		L"Is the tracer tritium/initial tritium - True or False",
		L"Dissolved inorganic carbon concentration of 1st component - number",
		L"Dissolved inorganic carbon concentration of 2ndFrac - number"
	},
	{ L"BMM_EMM",
		L"BKKBBBBBBMMMBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"BMM_EMM",
		L"Date range, Tracer input range, Age of first component, Sample date, Lambda, UZtt, MixFrac, Conc2ndComp, Is3He(trit), Is3H, Is3H/3Ho, DIC1, DIC2",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet tracer concentration from a binary mixture with an exponential mixing model for the 1st component",   
		L"The range corresponding to the tracer input dates (in decimal years).",                   
		L"The range corresponding to the tracer input concentration data.",
		L"The mean age of the 1st component in years - number",
		L"The sample date in decimal years - number",
		L"The decay rate of the tracer [ln(2)/half-life] - number",
		L"The unsaturated zone travel time in years - number",
		L"Mixing fraction of the 1st component - number between 0 and 1",
		L"Concentration for the 2nd component or formula for 2nd component - number",
		L"Is the tracer tritiogenic helium-3 - True or False",
		L"Is the tracer initial tritium - True or False",
		L"Is the tracer tritium/initial tritium - True or False",  
		L"Dissolved inorganic carbon concentration of 1st component - number",
		L"Dissolved inorganic carbon concentration of 2ndFrac - number"   
	},
	{ L"BMM_PEM",
		L"BKKBBBBBBBBMMMBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"BMM_PEM",
		L"Date range, Tracer input range, Tau, Sample date, Lambda, PEM Upper ratio, PEM Lower ratio, UZtt, MixFrac, Conc2ndComp, Is3He(trit), Is3Ho, Is3H/3Ho, DIC1, DIC2",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet tracer concentration from a binary mixture with a partial exponential model for the 1st component",   
		L"The range corresponding to the tracer input dates (in decimal years).",                   
		L"The range corresponding to the tracer input concentration data.",
		L"The mean age of the 1st component in years - number",
		L"The sample date in decimal years - number",
		L"The decay rate of the tracer [ln(2)/half-life] - number",
		L"The ratio of the distance between the water table and top of the screen and the distance between the top of the screen and the bottom of the aquifer - number",
		L"The ratio of the distance between the water table and bottom of the screen and the distance between the bottom of the screen and the bottom of the aquifer - number",
		L"The unsaturated zone travel time in years - number",
		L"Mixing fraction of the 1st component - number between 0 and 1",
		L"Concentration for the 2nd component or formula for 2nd component - number",
		L"Is the tracer tritiogenic helium-3 - True or False",
		L"Is the tracer initial tritium - True or False",
		L"Is the tracer tritium/initial tritium - True or False",  
		L"Dissolved inorganic carbon concentration of 1st component - number",
		L"Dissolved inorganic carbon concentration of 2ndFrac - number"   
	},
	{ L"BMM_EPM",
		L"BKKBBBBBBBMMMBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"BMM_EPM",
		L"Date range, Tracer input range, Tau, Sample date, Lambda, EPM ratio, UZtt, MixFrac, Conc2ndComp, Is3He(trit), Is3H, Is3H/3Ho, DIC1, DIC2",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet tracer concentration from a binary mixture with an exponential piston-flow model for the 1st component",   
		L"The range corresponding to the tracer input dates (in decimal years).",                   
		L"The range corresponding to the tracer input concentration data.",
		L"The mean age of the 1st component in years - number",
		L"The sample date in decimal years - number",
		L"The decay rate of the tracer [ln(2)/half-life] - number",
		L"The ratio of the recharge length to the length of piston-flow travel - number",
		L"The unsaturated zone travel time in years - number",
		L"Mixing fraction of the 1st component - number between 0 and 1",
		L"Concentration for the 2nd component or formula for 2nd component - number",
		L"Is the tracer tritiogenic helium-3 - True or False",
		L"Is the tracer initial tritium - True or False",
		L"Is the tracer tritium/initial tritium - True or False",  
		L"Dissolved inorganic carbon concentration of 1st component - number",
		L"Dissolved inorganic carbon concentration of 2ndFrac - number" 
	},
	{ L"BMM_GAM",
		L"BKKBBBBBBBAAABB$",                   // up to 255 args in Excel 2007,  
										   // upto 29 args in Excel 2003 and earlier versions
		L"BMM_GAM",
		L"Date range, Tracer input range, Age of young component, Sample date, Decay rate (lambda), Alpha, Unsat. zone travel time, Mixing fraction (young part), Concentration of old component, Is helium-3, Is initial tritium, Is tritium/initial tritium ratio, DIC of young component, DIC of old component",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet tracer concentration from a binary mixture with an exponential piston-flow model for the young fraction",   
		L"The range corresponding to the tracer input dates (in decimal years).",                   
		L"The range corresponding to the tracer input concentration data.",
		L"The mean age in years - number",
		L"The sample date in decimal years - number",
		L"The decay rate of the tracer [ln(2)/half-life] - number",
		L"The shape parameter of the GAMMA function",
		L"The unsaturated zone travel time in years - number",
		L"Mixing fraction of the young fraction - number between 0 and 1",
		L"Concentration of old fraction or formula for old concentration - number",
		L"Is the tracer helium-3 - True or False",
		L"Is the tracer initial tritium - True or False",
		L"Is the tracer tritium/initial tritium - True or False",  
		L"Dissolved inorganic carbon concentration of young fraction - number",
		L"Dissolved inorganic carbon concentration of old fraction - number" 
	},
	{ L"BMM_PFM",
		L"BKKBBBBBBMMMBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"BMM_PFM",
		L"Date range, Tracer input range, Tau, Sample date, Lambda, UZtt, MixFrac, Conc2ndComp, Is3He(trit), Is3H, Is3H/3Ho, DIC1, DIC2",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet tracer concentration from a binary mixture with a piston-flow model for the 1st component",   
		L"The range corresponding to the tracer input dates (in decimal years).",                   
		L"The range corresponding to the tracer input concentration data.",
		L"The mean age of the 1st component in years - number",
		L"The sample date in decimal years - number",
		L"The decay rate of the tracer [ln(2)/half-life] - number",
		L"The unsaturated zone travel time in years - number",
		L"Mixing fraction of the 1st component - number between 0 and 1",
		L"Concentration for the 2nd component or formula for 2nd component - number",
		L"Is the tracer tritiogenic helium-3 - True or False",
		L"Is the tracer initial tritium - True or False",
		L"Is the tracer tritium/initial tritium - True or False",  
		L"Dissolved inorganic carbon concentration of 1st component - number",
		L"Dissolved inorganic carbon concentration of 2ndFrac - number"   
	},
	{ L"BMM_DM_He4",
		L"BBBBBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"BMM_DM_He4",
		L"Uranium, Thorium, Phi, Rho, Tau, DP, HeSolRate, MixFrac, Conc2ndComp",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet Helium-4 concentration from a binary mixture with a disperion model for the 1st component",
		L"Uranium concentration of aquifer in parts per million - number",                   
		L"Thorium concentration of aquifer in parts per million - number",
		L"Porosity - number",
		L"Bulk density - number",
		L"The mean age of the 1st component in years - number",
		L"The ratio of dispersion to advection or inverse of the Peclet number (D/vx) - number",
		L"Helium solution rate (cc@STP/gramH2O/year) - number",
		L"Mixing fraction of the 1st component - number between 0 and 1",
		L"Concentration for the 2nd component or formula for 2nd component - number"
	},
	{ L"BMM_EMM_He4",
		L"BBBBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"BMM_EMM_He4",
		L"Uranium, Thorium, Phi, Rho, Tau, HeSolRate, MixFrac, Conc2ndComp",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet Helium-4 concentration from a binary mixture with an exponential mixing model for the 1st component",
		L"Uranium concentration of aquifer in parts per million - number",                   
		L"Thorium concentration of aquifer in parts per million - number",
		L"Porosity - number",
		L"Bulk density - number",
		L"The mean age of the 1st component in years - number",
		L"Helium solution rate (cc@STP/gramH2O/year) - number",
		L"Mixing fraction of the 1st component - number between 0 and 1",
		L"Concentration for the 2nd component or formula for 2nd component - number"   
	},
	{ L"BMM_PEM_He4",
		L"BBBBBBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"BMM_PEM_He4",
		L"Uranium, Thorium, Phi, Rho, Tau, PEM Upper ratio, PEM Lower ratio, HeSolRate, MixFrac, Conc2ndComp",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet Helium-4 concentration from a binary mixture with a partial exponential mixing model for the 1st component",
		L"Uranium concentration of aquifer in parts per million - number",                   
		L"Thorium concentration of aquifer in parts per million - number",
		L"Porosity - number",
		L"Bulk density - number",
		L"The mean age of the 1st component in years - number",
		L"The ratio of the distance between the water table and top of the screen and the distance between the top of the screen and the bottom of the aquifer - number",
		L"The ratio of the distance between the water table and bottom of the screen and the distance between the bottom of the screen and the bottom of the aquifer - number",
		L"Helium solution rate (cc@STP/gramH2O/year) - number",
		L"Mixing fraction of the 1st component - number between 0 and 1",
		L"Concentration for the 2nd component or formula for 2nd component - number"   
	},
	{ L"BMM_EPM_He4",
		L"BBBBBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"BMM_EPM_He4",
		L"Uranium, Thorium, Phi, Rho, Tau, EPM ratio, HeSolRate, MixFrac, Conc2ndComp",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet Helium-4 concentration from a binary mixture with an exponential mixing model for the 1st component",
		L"Uranium concentration of aquifer in parts per million - number",                   
		L"Thorium concentration of aquifer in parts per million - number",
		L"Porosity - number",
		L"Bulk density - number",
		L"The mean age of the 1st component in years - number",
		L"The ratio of the recharge length to the length of piston-flow travel - number",
		L"Helium solution rate (cc@STP/gramH2O/year) - number",
		L"Mixing fraction of the 1st component - number between 0 and 1",
		L"Concentration for the 2nd component or formula for 2nd component - number"   
	},
	{ L"BMM_GAM_He4",
		L"BBBBBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"BMM_GAM_He4",
		L"Uranium ppm, Thorium ppm, Porosity (phi), Sediment density (rho), Age of young component, Alpha, Helium solution rate, Mixing fraction (young part), Concentration of old component",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet Helium-4 concentration from a binary mixture with an exponential mixing model for the young fraction",
		L"Uranium concentration of aquifer in parts per million - number",                   
		L"Thorium concentration of aquifer in parts per million - number",
		L"Porosity - number",
		L"Sediment density - number",
		L"The mean age in years - number",
		L"The shape parameter of the GAMMA function - number",
		L"Helium solution rate (per year) - number",
		L"Mixing fraction of the young fraction - number between 0 and 1",
		L"Concentration of old fraction or formula for old concentration - number"   
	},
	{ L"BMM_PFM_He4",
		L"BBBBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"BMM_PFM_He4",
		L"Uranium, Thorium, Phi, Rho, Tau, HeSolRate, MixFrac, Conc2ndComp",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet Helium-4 concentration from a binary mixture with an exponential mixing model for the 1st component",
		L"Uranium concentration of aquifer in parts per million - number",                   
		L"Thorium concentration of aquifer in parts per million - number",
		L"Porosity - number",
		L"Bulk density - number",
		L"The mean age of the 1st component in years - number",
		L"Helium solution rate (cc@STP/gramH2O/year) - number",
		L"Mixing fraction of the 1st component - number between 0 and 1",
		L"Concentration for the 2nd component or formula for 2nd component - number"    
	},
	{ L"UserDefinedTracerOut",
		L"BKKKKBBBMMM$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"UserDefinedTracerOut",
		L"Date range, Tracer input range, Age range, Recharge fraction range, Sample date, Lambda, UZtt, Is3He(trit), Is3H, Is3H/3Ho",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet tracer concentration from a user-defined age distribution",
		L"The range corresponding to the tracer input dates (in decimal years)",                   
		L"The range corresponding to the tracer input concentration data",
		L"The range corresponding to the ages in a user-defined age distribution",                   
		L"The range corresponding to the fraction of recharge in a user-defined age distribution",
		L"The sample date in decimal years - number",
		L"The decay rate of the tracer [ln(2)/half-life] - number",
		L"The unsaturated zone travel time in years - number",
		L"Is the tracer tritiogenic helium-3 - True or False",
		L"Is the tracer initial tritium - True or False",
		L"Is the tracer tritium/initial tritium - True or False"   
	},
	{ L"UserDefinedTracerOut_He4",
		L"BBBBBKKB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"UserDefinedTracerOut_He4",
		L"Uranium, Thorium, Phi, Rho, Age range, Recharge fraction range, HeSolRate",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the outlet tracer concentration from a user-defined age distribution",
		L"Uranium concentration of aquifer in parts per million - number",                   
		L"Thorium concentration of aquifer in parts per million - number",
		L"Porosity - number",
		L"Bulk density - number",
		L"The range corresponding to the ages in a user-defined age distribution",                   
		L"The range corresponding to the fraction of recharge in a user-defined age distribution",
		L"Helium solution rate (cc@STP/gramH2O/year) - number"
	},
	{ L"gt_BMM_PFM",
		L"BBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"gt_BMM_PFM",
		L"Min Age, Max Age, Tau, UZtt, MixFrac, 2ndFrac",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the fractional contribution of flow for a given age interval"
		L" from a binary mixtue with a piston-flow model for the 1st component",   
		L"The minimum age of the age interval - number",  
		L"The maximum age of the age interval - number",   
		L"The mean age of the 1st component in years - number",
		L"The unsaturated zone travel time in years - number",
		L"Mixing fraction of the 1st component - number between 0 and 1",
		L"Fraction of recharge in 2nd component or formula for 2nd component - number"
	},
	{ L"gt_BMM_EMM",
		L"BBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"gt_BMM_EMM",
		L"Min Age, Max Age, Tau, UZtt, MixFrac, 2ndFrac",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the fractional contribution of flow for a given age interval"
		L" from a binary mixtue with a exponential mixing model for the 1st component",   
		L"The minimum age of the age interval - number",  
		L"The maximum age of the age interval - number",   
		L"The mean age of the 1st component in years - number",
		L"The unsaturated zone travel time in years - number",
		L"Mixing fraction of the 1st component - number between 0 and 1",
		L"Fraction of recharge in 2nd component or formula for 2nd component - number"
	},
	{ L"gt_BMM_PEM",
		L"BBBBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"gt_BMM_PEM",
		L"Min Age, Max Age, Tau, PEM Upper ratio, PEM Lower ratio, UZtt, MixFrac, 2ndFrac",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the fractional contribution of flow for a given age interval"
		L" from a binary mixtue with a partial exponential model for the 1st component",   
		L"The minimum age of the age interval - number",  
		L"The maximum age of the age interval - number",   
		L"The mean age of the 1st component in years - number",
		L"The ratio of the distance between the water table and top of the screen and the distance between the top of the screen and the bottom of the aquifer - number",
		L"The ratio of the distance between the water table and bottom of the screen and the distance between the bottom of the screen and the bottom of the aquifer - number",
		L"The unsaturated zone travel time in years - number",
		L"Mixing fraction of the 1st component - number between 0 and 1",
		L"Fraction of recharge in 2nd component or formula for 2nd component - number"
	},
	{ L"gt_BMM_EPM",
		L"BBBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"gt_BMM_EPM",
		L"Min Age, Max Age, Tau, EPM ratio, UZtt, MixFrac, 2ndFrac",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the fractional contribution of flow for a given age interval"
		L" from a binary mixtue with a exponential piston-flow mixing model for the 1st component",   
		L"The minimum age of the age interval - number",  
		L"The maximum age of the age interval - number",   
		L"The mean age of the 1st component in years - number",
		L"The ratio of the recharge length to the length of piston-flow travel - number",
		L"The unsaturated zone travel time in years - number",
		L"Mixing fraction of the 1st component - number between 0 and 1",
		L"Fraction of recharge in 2nd component or formula for 2nd component - number"
	},
	{ L"gt_BMM_GAM",
		L"BBBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"gt_BMM_GAM",
		L"Min Age, Max Age, Tau, Alpha, UZtt, MixFrac, 2ndFrac",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the fractional contribution of flow for a given age interval"
		L" from a binary mixtue with a exponential piston-flow mixing model for the 1st component",   
		L"The minimum age of the age interval - number",  
		L"The maximum age of the age interval - number",   
		L"The mean age of the 1st component in years - number",
		L"Alpha - number",
		L"The unsaturated zone travel time in years - number",
		L"Mixing fraction of the 1st component - number between 0 and 1",
		L"Fraction of recharge in 2nd component or formula for 2nd component - number"
	},
	{ L"gt_BMM_DM",
		L"BBBBBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"gt_BMM_DM",
		L"Min Age, Max Age, Tau, DP, UZtt, MixFrac, 2ndFrac",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the fractional contribution of flow for a given age interval"
		L" from a binary mixtue with a dispersion model for the 1st component",   
		L"The minimum age of the age interval - number",  
		L"The maximum age of the age interval - number",   
		L"The mean age of the 1st component in years - number",
		L"The ratio of dispersion to advection or inverse of the Peclet number (D/vx) - number",
		L"The unsaturated zone travel time in years - number",
		L"Mixing fraction of the 1st component - number between 0 and 1",
		L"Fraction of recharge in 2nd component or formula for 2nd component - number"
	},
	{ L"gt_UserDefinedAge",
		L"BKKBBBB$",                   // up to 255 args in Excel 2007, 
										   // upto 29 args in Excel 2003 and earlier versions
		L"gt_UserDefinedAge",
		L"Age range, Recharge fraction range, Year, Time increment, UZtt, MinAgeBin",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Returns the fractional contribution of flow for a given age interval of a user-defined age distribution",   
		L"The age range of the user-defined age distribution - range",  
		L"The recharge fraction range of the user-defined age distribution - range",   
		L"The age interval to calculate the fraction of recharge over - number",
		L"The increment of time between age intervals - number",
		L"The unsaturated zone travel time in years - number",                  
		L"The minimum difference between age bins in the user defined age distribution - number"
	}
};
//
// Later, the instance handle is required to create dialog boxes.
// g_hInst holds the instance handle passed in by DllMain so that it is
// available for later use. hWndMain is used in several routines to
// store Microsoft Excel's hWnd. This is used to attach dialog boxes as
// children of Microsoft Excel's main window. A buffer is used to store
// the free space that DIALOGMsgProc will put into the dialog box.
//
//
///***************************************************************************
// DllMain()
//
// Purpose:
//
//      Windows calls DllMain, for both initialization and termination.
//		It also makes calls on both a per-process and per-thread basis,
//		so several initialization calls can be made if a process is multithreaded.
//
//      This function is called when the DLL is first loaded, with a dwReason
//      of DLL_PROCESS_ATTACH.
//
// Parameters:
//
//      HANDLE hDLL         Module handle.
//      DWORD dwReason,     Reason for call
//      LPVOID lpReserved   Reserved
//
// Returns: 
//      The function returns TRUE (1) to indicate success. If, during
//      per-process initialization, the function returns zero, 
//      the system cancels the process.
//
// Comments:
//
// History:  Date       Author        Reason
///***************************************************************************

BOOL APIENTRY DllMain( HANDLE hDLL, 
					   DWORD dwReason, 
					   LPVOID lpReserved )
{
	switch (dwReason)
	{
	case DLL_PROCESS_ATTACH:

		// The instance handle passed into DllMain is saved
		// in the global variable g_hInst for later use.

		g_hInst = hDLL;
		break;
	case DLL_PROCESS_DETACH:
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	default:
		break;
	}
	return TRUE;
}


///***************************************************************************
// xlAutoOpen()
//
// Purpose: 
//      Microsoft Excel call this function when the DLL is loaded.
//
//      Microsoft Excel uses xlAutoOpen to load XLL files.
//      When you open an XLL file, the only action
//      Microsoft Excel takes is to call the xlAutoOpen function.
//
//      More specifically, xlAutoOpen is called:
//
//       - when you open this XLL file from the File menu,
//       - when this XLL is in the XLSTART directory, and is
//         automatically opened when Microsoft Excel starts,
//       - when Microsoft Excel opens this XLL for any other reason, or
//       - when a macro calls REGISTER(), with only one argument, which is the
//         name of this XLL.
//
//      xlAutoOpen is also called by the Add-in Manager when you add this XLL 
//      as an add-in. The Add-in Manager first calls xlAutoAdd, then calls
//      REGISTER("EXAMPLE.XLL"), which in turn calls xlAutoOpen.
//
//      xlAutoOpen should:
//
//       - register all the functions you want to make available while this
//         XLL is open,
//
//       - add any menus or menu items that this XLL supports,
//
//       - perform any other initialization you need, and
//
//       - return 1 if successful, or return 0 if your XLL cannot be opened.
//
// Parameters:
//
// Returns: 
//
//      int         1 on success, 0 on failure
//
// Comments:
//
// History:  Date       Author        Reason
///***************************************************************************

__declspec(dllexport) int WINAPI xlAutoOpen(void)
{

	static XLOPER12 xDLL,	   // name of this DLL //
	xMenu,	 // xltypeMulti containing the menu //
	xTool,	 // xltypeMulti containing the toolbar //
	xTest;	 // used for menu test //
	int i;			   // Loop indices //
	//
	// In the following block of code the name of the XLL is obtained by
	// calling xlGetName. This name is used as the first argument to the
	// REGISTER function to specify the name of the XLL. Next, the XLL loops
	// through the g_rgWorksheetFuncs[] table, and the g_rgCommandFuncs[]
	// tableregistering each function in the table using xlfRegister. 
	// Functions must be registered before you can add a menu item.
	//
	
	Excel12f(xlGetName, &xDLL, 0);

	for (i=0; i<g_rgWorksheetFuncsRows; i++)
	{
		Excel12f(xlfRegister, 0, 1+ g_rgWorksheetFuncsCols,
			  (LPXLOPER12) &xDLL,
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][0]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][1]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][2]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][3]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][4]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][5]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][6]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][7]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][8]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][9]), 
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][10]), 
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][11]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][12]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][13]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][14]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][15]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][16]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][17]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][18]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][19]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][20]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][21]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][22]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][23]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][24]));
	}
	
	// Free the XLL filename //
	Excel12f(xlFree, 0, 2, (LPXLOPER12) &xTest, (LPXLOPER12) &xDLL);

	return 1;
}


///***************************************************************************
// xlAutoClose()
//
// Purpose: Microsoft Excel call this function when the DLL is unloaded.
//
//      xlAutoClose is called by Microsoft Excel:
//
//       - when you quit Microsoft Excel, or 
//       - when a macro sheet calls UNREGISTER(), giving a string argument
//         which is the name of this XLL.
//
//      xlAutoClose is called by the Add-in Manager when you remove this XLL from
//      the list of loaded add-ins. The Add-in Manager first calls xlAutoRemove,
//      then calls UNREGISTER("GENERIC.XLL"), which in turn calls xlAutoClose.
// 
//      xlAutoClose is called by GENERIC.XLL by the function fExit. This function
//      is called when you exit Generic.
// 
//      xlAutoClose should:
// 
//       - Remove any menus or menu items that were added in xlAutoOpen,
// 
//       - do any necessary global cleanup, and
// 
//       - delete any names that were added (names of exported functions, and 
//         so on). Remember that registering functions may cause names to 
//         be created.
// 
//      xlAutoClose does NOT have to unregister the functions that were registered
//      in xlAutoOpen. This is done automatically by Microsoft Excel after
//      xlAutoClose returns.
// 
//      xlAutoClose should return 1.
//
// Parameters:
//
// Returns: 
//
//      int         1
//
// Comments:
//
// History:  Date       Author        Reason
///***************************************************************************

__declspec(dllexport) int WINAPI xlAutoClose(void)
{
	int i;
	//
	// This block first deletes all names added by xlAutoOpen or
	// xlAutoRegister12. Next, it checks if the drop-down menu Generic still
	// exists. If it does, it is deleted using xlfDeleteMenu. It then checks
	// if the Test toolbar still exists. If it is, xlfDeleteToolbar is
	// used to delete it.
	//

	//
	// Due to a bug in Excel the following code to delete the defined names
	// does not work.  There is no way to delete these
	// names once they are Registered
	// The code is left in, in hopes that it will be
	// fixed in a future version.
	//

	for (i = 0; i < g_rgWorksheetFuncsRows; i++)
	{
		Excel12f(xlfSetName, 0, 1, TempStr12(g_rgWorksheetFuncs[i][2]));
	}

	return 1;
}


///***************************************************************************
// lpwstricmp()
//
// Purpose: 
//
//      Compares a pascal string and a null-terminated C-string to see
//      if they are equal.  Method is case insensitive
//
// Parameters:
//
//      LPWSTR s    First string (null-terminated)
//      LPWSTR t    Second string (byte counted)
//
// Returns: 
//
//      int         0 if they are equal
//                  Nonzero otherwise
//
// Comments:
//
//      Unlike the usual string functions, lpwstricmp
//      doesn't care about collating sequence.
//
// History:  Date       Author        Reason
///***************************************************************************

int lpwstricmp(LPWSTR s, LPWSTR t)
{
	int i;

	if (wcslen(s) != *t)
		return 1;

	for (i = 1; i <= s[0]; i++)
	{
		if (towlower(s[i-1]) != towlower(t[i]))
			return 1;
	}										  
	return 0;
}


///***************************************************************************
// xlAutoRegister12()
//
// Purpose:
//
//      This function is called by Microsoft Excel if a macro sheet tries to
//      register a function without specifying the type_text argument. If that
//      happens, Microsoft Excel calls xlAutoRegister12, passing the name of the
//      function that the user tried to register. xlAutoRegister12 should use the
//      normal REGISTER function to register the function, only this time it must
//      specify the type_text argument. If xlAutoRegister12 does not recognize the
//      function name, it should return a #VALUE! error. Otherwise, it should
//      return whatever REGISTER returned.
//
// Parameters:
//
//      LPXLOPER12 pxName   xltypeStr containing the
//                          name of the function
//                          to be registered. This is not
//                          case sensitive.
//
// Returns: 
//
//      LPXLOPER12          xltypeNum containing the result
//                          of registering the function,
//                          or xltypeErr containing #VALUE!
//                          if the function could not be
//                          registered.
//
// Comments:
//
// History:  Date       Author        Reason
///***************************************************************************

__declspec(dllexport) LPXLOPER12 WINAPI xlAutoRegister12(LPXLOPER12 pxName)
{
	static XLOPER12 xDLL, xRegId;
	int i;

	//
	// This block initializes xRegId to a #VALUE! error first. This is done in
	// case a function is not found to register. Next, the code loops through 
	// the functions in g_rgFuncs[] and uses lpwstricmp to determine if the 
	// current row in g_rgFuncs[] represents the function that needs to be 
	// registered. When it finds the proper row, the function is registered 
	// and the register ID is returned to Microsoft Excel. If no matching 
	// function is found, an xRegId is returned containing a #VALUE! error.
	//

	xRegId.xltype = xltypeErr;
	xRegId.val.err = xlerrValue;


	for (i=0; i<g_rgWorksheetFuncsRows; i++)
	{
		if (!lpwstricmp(g_rgWorksheetFuncs[i][0], pxName->val.str))
		{
			Excel12f(xlfRegister, 0, 1+ g_rgWorksheetFuncsCols,
				  (LPXLOPER12) &xDLL,
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][0]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][1]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][2]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][3]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][4]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][5]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][6]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][7]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][8]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][9]), 
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][10]), 
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][11]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][12]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][13]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][14]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][15]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][16]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][17]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][18]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][19]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][20]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][21]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][22]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][23]));
			/// Free oper returned by xl //
			Excel12f(xlFree, 0, 1, (LPXLOPER12) &xDLL);

			return(LPXLOPER12) &xRegId;
		}
	}
	return 0;
}

///***************************************************************************
// xlAutoAdd()
//
// Purpose:
//
//      This function is called by the Add-in Manager only. When you add a
//      DLL to the list of active add-ins, the Add-in Manager calls xlAutoAdd()
//      and then opens the XLL, which in turn calls xlAutoOpen.
//
// Parameters:
//
// Returns: 
//
//      int         1
//
// Comments:
//
// History:  Date       Author        Reason
///***************************************************************************

__declspec(dllexport) int WINAPI xlAutoAdd(void)
{
	//XCHAR szBuf[255];

	//wsprintfW((LPWSTR)szBuf, L"Thank you for adding TracerLPMfunctions.XLL\n "
	//		 L"built on %hs at %hs", __DATE__, __TIME__);

	// Display a dialog box indicating that the XLL was successfully added //
	//Excel12f(xlcAlert, 0, 2, TempStr12(szBuf), TempInt12(2));
	return 1;
}

///***************************************************************************
// xlAutoRemove()
//
// Purpose:
//
//      This function is called by the Add-in Manager only. When you remove
//      an XLL from the list of active add-ins, the Add-in Manager calls
//      xlAutoRemove() and then UNREGISTER("GENERIC.XLL").
//   
//      You can use this function to perform any special tasks that need to be
//      performed when you remove the XLL from the Add-in Manager's list
//      of active add-ins. For example, you may want to delete an
//      initialization file when the XLL is removed from the list.
//
// Parameters:
//
// Returns: 
//
//      int         1
//
// Comments:
//
// History:  Date       Author        Reason
///***************************************************************************

__declspec(dllexport) int WINAPI xlAutoRemove(void)
{
	// Show a dialog box indicating that the XLL was successfully removed //
	//Excel12f(xlcAlert, 0, 2, TempStr12(L"Thank you for removing TracerLPMfunctions.XLL!"),
	//	  TempInt12(2));
	return 1;
}

///***************************************************************************
// xlAddInManagerInfo12()
//
// Purpose:
//
//      This function is called by the Add-in Manager to find the long name
//      of the add-in. If xAction = 1, this function should return a string
//      containing the long name of this XLL, which the Add-in Manager will use
//      to describe this XLL. If xAction = 2 or 3, this function should return
//      #VALUE!.
//
// Parameters:
//
//      LPXLOPER12 xAction  What information you want. One of:
//                            1 = the long name of the
//                                add-in
//                            2 = reserved
//                            3 = reserved
//
// Returns: 
//
//      LPXLOPER12          The long name or #VALUE!.
//
// Comments:
//
// History:  Date       Author        Reason
///***************************************************************************

__declspec(dllexport) LPXLOPER12 WINAPI xlAddInManagerInfo12(LPXLOPER12 xAction)
{
	static XLOPER12 xInfo, xIntAction;

	//
	// This code coerces the passed-in value to an integer. This is how the
	// code determines what is being requested. If it receives a 1, 
	// it returns a string representing the long name. If it receives 
	// anything else, it returns a #VALUE! error.
	//

	Excel12f(xlCoerce, &xIntAction, 2, xAction, TempInt12(xltypeInt));

	if (xIntAction.val.w == 1)
	{
		xInfo.xltype = xltypeStr;
		xInfo.val.str = L"\023TracerLPM functions DLL";
	}
	else
	{
		xInfo.xltype = xltypeErr;
		xInfo.val.err = xlerrValue;
	}

	//Word of caution - returning static XLOPERs/XLOPER12s is not thread safe
	//for UDFs declared as thread safe, use alternate memory allocation mechanisms
	return(LPXLOPER12) &xInfo;
}

static const double Tol= 1E-06;        // Stopping criteria for LPM output 
static const double Udecay = 1.19E-13;
static const double THdecay = 2.88E-14;
static const double PI = 3.1415926535897932384626433832795028841971693993751;
static const double RndNum[5] = {0.0521,0.2311,0.00125,0.4860,0.8913};
static const double glX[3] = {0.94288241569547971905635175843185720232,0.64185334234578130578123554132903188354,0.23638319966214988028222377349205292599};
static const double glA[7] = {0.015827191973480183087169986733305510591,0.094273840218850045531282505077108171960,0.15507198733658539625363597980210298680,
				0.18882157396018245442000533937297167125, 0.19977340522685852679206802206648840246, 0.22492646533333952701601768799639508076, 0.24261107190140773379964095790325635233};
static const double Roots5[5] = {0.906179845938664,0.538469310105683,0.0,-0.538469310105683,-0.906179845938664};
static const double Coeff5[5] = {0.236926885056189,0.478628670499366,0.568888888888889,0.478628670499366,0.236926885056189};
static const double Roots3[3] = {sqrt(0.6),0.0,-sqrt(0.6)};
static const double Coeff3[3] = {double(5)/double(9),double(8)/double(9),double(5)/double(9)};

int sign(const double X)
{   
	return ((X==0)?0:(X>0)?1:-1);
}

double ReturnLambdaCorrection(double MinAge, double MaxAge, double Lambda)
{
	return double (exp(-Lambda*(MinAge+(MaxAge-MinAge)/2)));
}

double MinTimeInc(FP DateRange[])
{
	double Result = 1E06;
	__int32 i;
	for (i = 1; i<= DateRange->rows; i++)
	{
		if (fabs(DateRange->array[i]-DateRange->array[i-1])<Result)
		{
			Result = DateRange->array[i]-DateRange->array[i-1];
		}
	}
	return Result;
}

__declspec(dllexport) double WINAPI gt_UserDefinedAge(FP AgeRange[], FP RechFracRange[], double Year, double TimeInc, double UZtime, double MinAgeBin)
{

	__int32 i=0, Rows=0;
	__int32 iCount=0;
	double CurAge, FracSum=0.0, Result=0.0;

	if ((AgeRange->rows == RechFracRange->rows) && (MinAgeBin != 0.0) && (Year > 0.0))
	{
		Rows = RechFracRange->rows - 1;
		Rows = AgeRange->rows - 1;
		CurAge = Year + UZtime;
		while ((AgeRange->array[i] <= CurAge - TimeInc) && (i < Rows))
		{
			i++;
		}
		if (TimeInc < MinAgeBin && AgeRange->array[i] - MinAgeBin < CurAge && CurAge <= AgeRange->array[Rows])
		{
			FracSum = RechFracRange->array[i];
			Result = FracSum * TimeInc / MinAgeBin;
			return Result;
		}
		else if (TimeInc >= MinAgeBin && CurAge - TimeInc < AgeRange->array[Rows])
		{
			while (CurAge >= AgeRange->array[i + iCount] && (i + iCount <= Rows))
			{
				FracSum += RechFracRange->array[i + iCount];
				iCount++;
			}
			Result = FracSum;
			return Result;
		}
	}
	return Result;
}
__declspec(dllexport) double WINAPI UserDefinedTracerOut(
                        FP DateRange[], FP TracerRange[], FP AgeRange[], FP RechFracRange[], double SampleDate, double Lambda, double UZtime,
						LPXLOPER12 HeliumThree, LPXLOPER12 InitialTrit, LPXLOPER12 TritInitialTritRatio)
{
	__int32 i=0, n=0;
	__int32 j, StepInc, StopCriteria;
	double DR, Cin, Result=0.0, Lambda2, Res2=0.0;
	
	j = DateRange->rows - 1;
	StepInc = 1;
	StopCriteria = 0;
	if (AgeRange->rows = RechFracRange->rows)
	{
		if (DateRange->array[j] < DateRange->array[0])
		{
			StepInc = -1;
			StopCriteria = j;
			j = 0;
		}
		DR = DateRange->array[j];
		while ((SampleDate - DR - UZtime <= AgeRange->array[i]) && j != StopCriteria)
		{
			j = j - StepInc;
			DR = DateRange->array[j];
		}
		if (HeliumThree->val.xbool == TRUE)
		{
			if (Lambda == 0.0)
			{
				Lambda2 = log (2) / 12.32;
			}
			else
			{
				Lambda2 = Lambda;
			}
			for (i = 0; i<= AgeRange->rows-2; i++)
			{
				Cin = TracerRange->array[j];
				n = 0;
				if (j != StopCriteria)
				{
					while ((SampleDate - DR - UZtime <= AgeRange->array[i+1]) && j != StopCriteria)
					{
						n++;
						j = j - StepInc;
						DR = DateRange->array[j];
						Cin += TracerRange->array[j];
					}
				}
				if (n > 0) 
				{
					Cin = (Cin - TracerRange->array[j]) / n;
				}
				Result += RechFracRange->array[i] * Cin * exp (-Lambda2 * UZtime) * (1 - exp (-Lambda2 * AgeRange->array[i]));
			}
			return Result;
		}
		if (TritInitialTritRatio->val.xbool == TRUE)
		{
			if (Lambda == 0.0)
			{
				Lambda2 = log (2) / 12.32;
			}
			else
			{
				Lambda2 = Lambda;
			}
			for (i = 0; i<= AgeRange->rows-2; i++)
			{
				Cin = TracerRange->array[j];
				n = 0;
				if (j != StopCriteria)
				{
					while ((SampleDate - DR - UZtime <= AgeRange->array[i+1]) && j != StopCriteria)
					{
						n++;
						j = j - StepInc;
						DR = DateRange->array[j];
						Cin += TracerRange->array[j];
					}
				}
				if (n > 0) 
				{
					Cin = (Cin - TracerRange->array[j]) / n;
				}
				Result += RechFracRange->array[i] * Cin * exp (-Lambda2 * UZtime) * exp (-Lambda * AgeRange->array[i]); //Tritium
				Res2 += RechFracRange->array[i] * Cin * exp (-Lambda2 * UZtime); // Initialtritium
			}
			Result = Result / Res2;
			return Result;
		}
		if (TritInitialTritRatio->val.xbool != TRUE && HeliumThree->val.xbool != TRUE)
		{
			if (InitialTrit->val.xbool == TRUE)
			{
				Lambda2 = log (2) / 12.32;
			}
			else
			{
				Lambda2 = Lambda;
			}
			for (i = 0; i<= AgeRange->rows-2; i++)
			{
				Cin = TracerRange->array[j];
				n = 0;
				if (j != StopCriteria)
				{
					while ((SampleDate - DR - UZtime <= AgeRange->array[i+1]) && j != StopCriteria)
					{
						n++;
						j = j - StepInc;
						DR = DateRange->array[j];
						Cin += TracerRange->array[j];
					}
				}
				if (n > 0) 
				{
					Cin = (Cin - TracerRange->array[j]) / n;
				}
				Result += RechFracRange->array[i] * Cin * exp (-Lambda2 * UZtime) * exp (-Lambda * AgeRange->array[i]);
			}
			return Result;
		}
	}
	return Result;
}

__declspec(dllexport) double WINAPI UserDefinedTracerOut_He4(double Uppm, double THppm, double Porosity, double SedDensity, FP AgeRange[], FP RechFracRange[], double HeSolnRate)
{
	double Result;
	__int32 i;
    Result = 0.0;
	if (AgeRange->rows = RechFracRange->rows)
	{	
		if (HeSolnRate != 0.0)
		{
			for (i = 0; i<= AgeRange->rows-1; i++)
			{
				Result += RechFracRange->array[i] * HeSolnRate * AgeRange->array[i];
			}
		}
		else
		{
			for (i = 0; i<= AgeRange->rows-1; i++)
			{
				Result += RechFracRange->array[i] * SedDensity * (Udecay * Uppm + THdecay * THppm) * AgeRange->array[i] / Porosity;
			}
		}
	}
	return Result;
}
__declspec(dllexport) double WINAPI PFM(
                        FP DateRange[], FP TracerRange[], double Tau, double SampleDate, 
						double Lambda, double UZtime, LPXLOPER12 HeliumThree, LPXLOPER12 InitialTrit, LPXLOPER12 TritInitialTritRatio)
{
	double DR, Cin, Result, Lambda2;
	__int32 j, StepInc, StopCriteria;
	Result = 0;
	j = DateRange->rows - 1;
	StepInc = 1;
	StopCriteria = 0;
	if (DateRange->array[j] < DateRange->array[0])
	{
		StepInc = -1;
		StopCriteria = j;
		j = 0;
	}
	DR = DateRange->array[j];
	if (HeliumThree->val.xbool == TRUE)
	{
		if (Lambda == 0.0)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		if (SampleDate - Tau - UZtime < DateRange->array[StopCriteria])
		{
			Result = TracerRange->array[StopCriteria] * exp(-Lambda2 * UZtime) * (1 - exp(-Lambda2 * Tau));
		}
		else
		{
			while ((SampleDate - Tau - UZtime < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
			}
			Cin = TracerRange->array[j];
			Result = Cin * exp(-Lambda2 * UZtime) * (1 - exp(-Lambda2 * Tau));
			return Result;
		}
	}
	if (TritInitialTritRatio->val.xbool == TRUE)
	{
		if (Lambda == 0.0)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		Result = exp(-Lambda2 * Tau);
		return Result;
	}
	if (TritInitialTritRatio->val.xbool != TRUE && HeliumThree->val.xbool != TRUE)
	{
		if (InitialTrit->val.xbool == TRUE)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		if (SampleDate - Tau - UZtime < DateRange->array[StopCriteria])
		{
			Result = TracerRange->array[StopCriteria] * exp(-Lambda2 * UZtime) * exp(-Lambda * Tau);
		}
		else
		{
			while ((SampleDate - Tau - UZtime < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
			}
			Cin = TracerRange->array[j];
			Result = Cin * exp(-Lambda2 * UZtime) * exp(-Lambda * Tau);
			return Result;
		}
	}
	return Result;
}

__declspec(dllexport) double WINAPI PFM_He4(double Uppm, double THppm, double Porosity, double SedDensity, double Tau, double HeSolnRate)
{
	double Result;
    Result = 0.0;
	if (HeSolnRate != 0.0)
	{
		Result = Tau * HeSolnRate;
	}
	else
	{
		Result = SedDensity / Porosity * Tau * (Udecay * Uppm + THdecay * THppm);
	}
	return Result;
}
__declspec(dllexport) double WINAPI EMM(FP DateRange[], FP TracerRange[], double Tau, double SampleDate, 
						double Lambda, double UZtime, LPXLOPER12 HeliumThree, LPXLOPER12 InitialTrit, LPXLOPER12 TritInitialTritRatio)
{
	double Result=0.0;
	double DR, Cin; // pointers to Date Range and Tracer Input
	double TauRes, n=1.0, EndDate, EMM1=0.0, CinHe3, EMMnoDecay=0.0, TimeIncrement, MaxDate;
	double EMMhalf1=0.0, EMMhalf2=0.0, Multiplier, MinAge, MaxAge, EMMnd1=0.0, EMMnd2=0.0, Lambda2;
	//bool TorF; // pointer to boolean values
	__int32 i,j, nIters, StepInc, StopCriteria;

	TauRes = Tau / (1 + log (n)); //Turnover Time of aquifer
	MinAge = SampleDate - log (n) * TauRes; //Age at Z star
	if (Tau > 100)
	{
		TimeIncrement = fabs(DateRange->array[1]-DateRange->array[0]); //MinTimeInc(DateRange);
		if (fabs(DateRange->array[DateRange->rows - 1]-DateRange->array[DateRange->rows - 2]) < TimeIncrement)
		{
			TimeIncrement = fabs(DateRange->array[DateRange->rows - 1]-DateRange->array[DateRange->rows - 2]);
		}
	}
	else
	{
		TimeIncrement = 1.0 / 12.0;
	}
	MaxAge = floor (MinAge) + (floor ((MinAge - floor (MinAge)) / TimeIncrement)) * TimeIncrement;
	EndDate = SampleDate - UZtime;
	MinAge = SampleDate - MinAge;
	MaxAge = SampleDate - MaxAge;
	nIters = 2000000;
	j = DateRange->rows - 1;
	StepInc = 1;
	StopCriteria = 0;
	if (DateRange->array[j] < DateRange->array[0])
	{
		StepInc = -1;
		StopCriteria = j;
		j = 0;
	}
	DR = DateRange->array[j];
	while ((DR >= EndDate - MinAge) && j != StopCriteria)
	{
		j = j - StepInc;
		DR = DateRange->array[j];
	}
	Cin = TracerRange->array[j];
	if (Cin == 0 && j == StopCriteria)
	{
		return Result;
	}
	//TorF = HeliumThree->val.num
	if (HeliumThree->val.xbool == TRUE)
	{
		if (Lambda == 0.0)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		Cin = TracerRange->array[j];
		CinHe3 = Cin * exp(-Lambda2 * UZtime) * (1 - exp(-Lambda2 * (MaxAge - (MaxAge-MinAge) / 2)));
		EMMhalf1 = exp(-(MinAge) / TauRes);
		EMMhalf2 = exp(-(MaxAge) / TauRes);
		Result = CinHe3 * n * (EMMhalf1 - EMMhalf2);
		MinAge = MaxAge;
		for (i = 1; i<= nIters; i++)
		{		
			EMMhalf1 = EMMhalf2;
			MaxAge = MinAge + i * TimeIncrement;
			while ((EndDate - MaxAge < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					return Result;
				}
			}
			EMM1 = Result;
			CinHe3 = Cin * exp(-Lambda2 * UZtime) * (1 - exp(-Lambda2 * (MaxAge - TimeIncrement / 2)));
			EMMhalf2 = exp(-(MaxAge) / TauRes);
			Result += CinHe3 * n * (EMMhalf1 - EMMhalf2);
			if (MaxAge >= Tau && Result > 0)
			{
				if ((Result - EMM1) / Result < Tol)
				{
					return Result;
				}
			}
		}
	}
	if (TritInitialTritRatio->val.xbool == TRUE)
	{
		if (Lambda == 0.0)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		Multiplier = n / TauRes * (1 / ((1 / TauRes) + Lambda2));
		EMMnd1 = exp(-(MinAge) / TauRes);
		EMMnd2 = exp(-(MaxAge) / TauRes);
		EMMnoDecay = TracerRange->array[j] * exp(-Lambda2 * UZtime) * n * (EMMnd1 - EMMnd2);
		EMMnd1 = EMMnd2;
		EMMhalf1 = exp(-(MinAge) * ((1 / TauRes) + Lambda2));
		EMMhalf2 = exp(-(MaxAge) * ((1 / TauRes) + Lambda2));
		Cin = TracerRange->array[j];
		Result = Cin * exp(-Lambda2 * UZtime) * Multiplier * (EMMhalf1 - EMMhalf2);
		MinAge = MaxAge;
		EMMhalf1 = EMMhalf2;
		for (i = 1; i<= nIters; i++)
		{
			EMM1 = Result;
			MaxAge = MinAge + i * TimeIncrement;
			while ((EndDate - MaxAge < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					break;
				}
			}
			EMMnd2 = exp(-MaxAge / TauRes);
			EMMnoDecay += Cin * exp(-Lambda2 * UZtime) * n * (EMMnd1 - EMMnd2);
			EMMnd1 = EMMnd2;
			EMMhalf2 = exp(-MaxAge * ((1 / TauRes) + Lambda2));
			Result += Cin * exp(-Lambda2 * UZtime) * Multiplier * (EMMhalf1 - EMMhalf2);
			EMMhalf1 = EMMhalf2;
			if (MaxAge >= Tau && Result > 0)
			{
				if ((Result - EMM1) / Result < Tol)
				{
					break;
				}
			}
		}
		Result = Result / EMMnoDecay;
		return Result;
	}
	if (TritInitialTritRatio->val.xbool != TRUE && HeliumThree->val.xbool != TRUE)
	{
		if (InitialTrit->val.xbool == TRUE)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		Multiplier = n / TauRes * (1 / ((1 / TauRes) + Lambda));
		EMMhalf1 = exp(-(MinAge) * ((1 / TauRes) + Lambda));
		EMMhalf2 = exp(-(MaxAge) * ((1 / TauRes) + Lambda));
		Cin = TracerRange->array[j];
		Result = Cin * exp(-Lambda2 * UZtime) * Multiplier * (EMMhalf1 - EMMhalf2);
		MinAge = MaxAge;
		for (i = 1; i<= nIters; i++)
		{
			EMMhalf1 = EMMhalf2;
			MaxAge = MinAge + i * TimeIncrement;
			MaxDate = EndDate - i * TimeIncrement;
			while (MaxDate < DR && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					return Result;
				}
			}
			EMM1 = Result;
			EMMhalf2 = exp(-(MaxAge) * ((1 / TauRes) + Lambda));
			Result += Cin * exp(-Lambda2 * UZtime) * Multiplier * (EMMhalf1 - EMMhalf2);
			if (MaxAge >= Tau && Result > 0)
			{
				if ((Result - EMM1) / Result < Tol)
				{
					return Result;
				}
			}
		}
	}
	return Result;
}

__declspec(dllexport) double WINAPI EMM_He4(
                        double Uppm, double THppm, double Porosity, double SedDensity, double Tau, double HeSolnRate)
{
	double Result;
	double Cin, EMMold, TimeIncrement, MinAge, MaxAge, n=1.0, TauRes, EMMhalf1, EMMhalf2;
	__int32 i, SimYears, nIters;
    
    TauRes = Tau / (1 + log (n));
    SimYears = 1500000;
    if (Tau < 100)
	{
		TimeIncrement = 1.0 / 12.0;
	}
	else
	{
		if (Tau < 1000)
		{
			TimeIncrement = 0.5;
		}
		else
		{
			TimeIncrement = Tau / 1000;
		}
	}
    MinAge = log (n) * TauRes; //Tstar
    MaxAge = floor (MinAge) + (floor ((MinAge-floor (MinAge)) / TimeIncrement) + 1) * TimeIncrement;
    nIters = (__int32) ((SimYears - MinAge) / TimeIncrement);
    Result = 0.0;
	if (HeSolnRate != 0.0)
	{
		Cin = HeSolnRate;
	}
	else
	{
		Cin = SedDensity / Porosity * (Udecay * Uppm + THdecay * THppm);
	}
    EMMhalf1 = exp(-MinAge / TauRes);
    EMMhalf2 = exp(-MaxAge / TauRes);
    Result = Cin * MaxAge * n * (EMMhalf1 - EMMhalf2);
	MinAge = MaxAge;
    for (i = 1; i<= nIters; i++)
	{
        EMMold = Result;
		EMMhalf1 = EMMhalf2;
        MaxAge = i * TimeIncrement + MinAge;
        EMMhalf2 = exp(-MaxAge / TauRes);
        Result += Cin * MaxAge * n * (EMMhalf1 - EMMhalf2);
        if (MaxAge >= Tau && Result > 0)
			{
				if ((Result - EMMold) / Result < Tol)
				{
					return Result;
				}
			}
	}
	return Result;
}

/*__declspec(dllexport) double WINAPI PEM(FP DateRange[], FP TracerRange[], double Tau, double SampleDate, 
						double Lambda, double PEMratio, double UZtime, LPXLOPER12 HeliumThree, LPXLOPER12 InitialTrit, LPXLOPER12 TritInitialTritRatio)
{
	double Result=0.0;
	double DR, Cin; // pointers to Date Range and Tracer Input
	double TauRes, n, EndDate, PEM1=0.0, CinHe3, PEMnoDecay=0.0, TimeIncrement;
	double PEMhalf1=0.0, PEMhalf2=0.0, Multiplier, MinAge, MaxAge, PEMnd1=0.0, PEMnd2=0.0, Lambda2;
	__int32 i,j, nIters, StepInc, StopCriteria;

	n = PEMratio + 1;
	TauRes = Tau / (1 + log (n)); //Turnover Time of aquifer
	MinAge = SampleDate - log (n) * TauRes; //Age at Z star
	if (Tau > 100)
	{
		TimeIncrement = fabs(DateRange->array[1]-DateRange->array[0]); //MinTimeInc(DateRange);
		if (fabs(DateRange->array[DateRange->rows - 1]-DateRange->array[DateRange->rows - 2]) < TimeIncrement)
		{
			TimeIncrement = fabs(DateRange->array[DateRange->rows - 1]-DateRange->array[DateRange->rows - 2]);
		}
	}
	else
	{
		TimeIncrement = 1.0 / 12.0;
	}
	MaxAge = floor (MinAge) + (floor ((MinAge - floor (MinAge)) / TimeIncrement)) * TimeIncrement;
	EndDate = SampleDate - UZtime;
	MinAge = SampleDate - MinAge;
	MaxAge = SampleDate - MaxAge;
	nIters = 2000000;
	j = DateRange->rows - 1;
	StepInc = 1;
	StopCriteria = 0;
	if (DateRange->array[j] < DateRange->array[0])
	{
		StepInc = -1;
		StopCriteria = j;
		j = 0;
	}
	DR = DateRange->array[j];
	while ((DR >= EndDate - MinAge) && j != StopCriteria)
	{
		j = j - StepInc;
		DR = DateRange->array[j];
	}
	Cin = TracerRange->array[j];
	if (Cin == 0 && j == StopCriteria)
	{
		return Result;
	}
	if (HeliumThree->val.xbool == TRUE)
	{
		if (Lambda == 0.0)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		CinHe3 = Cin * exp(-Lambda2 * UZtime) * (1 - exp(-Lambda2 * (MaxAge - (MaxAge-MinAge) / 2)));
		PEMhalf1 = exp(-(MinAge) / TauRes);
		PEMhalf2 = exp(-(MaxAge) / TauRes);
		Result = CinHe3 * n * (PEMhalf1 - PEMhalf2);
		MinAge = MaxAge;
		for (i = 1; i<= nIters; i++)
		{		
			PEMhalf1 = PEMhalf2;
			MaxAge = MinAge + i * TimeIncrement;
			while ((EndDate - MaxAge < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					return Result;
				}
			}
			PEM1 = Result;
			CinHe3 = Cin * exp(-Lambda2 * UZtime) * (1 - exp(-Lambda2 * (MaxAge - TimeIncrement / 2)));
			PEMhalf2 = exp(-(MaxAge) / TauRes);
			Result += CinHe3 * n * (PEMhalf1 - PEMhalf2);
			if (MaxAge >= Tau && Result > 0)
			{
				if ((Result - PEM1) / Result < Tol)
				{
					return Result;
				}
			}
		}
	}
	if (TritInitialTritRatio->val.xbool == TRUE)
	{
		if (Lambda == 0.0)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		Multiplier = n / TauRes * (1 / ((1 / TauRes) + Lambda2));
		PEMnd1 = exp(-(MinAge) / TauRes);
		PEMnd2 = exp(-(MaxAge) / TauRes);
		PEMnoDecay = TracerRange->array[j] * exp(-Lambda2 * UZtime) * n * (PEMnd1 - PEMnd2);
		PEMnd1 = PEMnd2;
		PEMhalf1 = exp(-(MinAge) * ((1 / TauRes) + Lambda2));
		PEMhalf2 = exp(-(MaxAge) * ((1 / TauRes) + Lambda2));
		Result = Cin * exp(-Lambda2 * UZtime) * Multiplier * (PEMhalf1 - PEMhalf2);
		MinAge = MaxAge;
		PEMhalf1 = PEMhalf2;
		for (i = 1; i<= nIters; i++)
		{
			PEM1 = Result;
			MaxAge = MinAge + i * TimeIncrement;
			while ((EndDate - MaxAge < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					break;
				}
			}
			PEMnd2 = exp(-MaxAge / TauRes);
			PEMnoDecay += Cin * exp(-Lambda2 * UZtime) * n * (PEMnd1 - PEMnd2);
			PEMnd1 = PEMnd2;
			PEMhalf2 = exp(-MaxAge * ((1 / TauRes) + Lambda2));
			Result += Cin * exp(-Lambda2 * UZtime) * Multiplier * (PEMhalf1 - PEMhalf2);
			PEMhalf1 = PEMhalf2;
			if (MaxAge >= Tau && Result > 0)
			{
				if ((Result - PEM1) / Result < Tol)
				{
					break;
				}
			}
		}
		Result = Result / PEMnoDecay;
		return Result;
	}
	if (TritInitialTritRatio->val.xbool != TRUE && HeliumThree->val.xbool != TRUE)
	{
		if (InitialTrit->val.xbool == TRUE)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		Multiplier = n / TauRes * (1 / ((1 / TauRes) + Lambda));
		PEMhalf1 = exp(-(MinAge) * ((1 / TauRes) + Lambda));
		PEMhalf2 = exp(-(MaxAge) * ((1 / TauRes) + Lambda));
		Result = Cin * exp(-Lambda2 * UZtime) * Multiplier * (PEMhalf1 - PEMhalf2);
		MinAge = MaxAge;
		for (i = 1; i<= nIters; i++)
		{
			PEMhalf1 = PEMhalf2;
			MaxAge = MinAge + i * TimeIncrement;
			while ((EndDate - MaxAge < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					return Result;
				}
			}
			PEM1 = Result;
			PEMhalf2 = exp(-(MaxAge) * ((1 / TauRes) + Lambda));
			Result += Cin * exp(-Lambda2 * UZtime) * Multiplier * (PEMhalf1 - PEMhalf2);
			if (MaxAge >= Tau && Result > 0)
			{
				if ((Result - PEM1) / Result < Tol)
				{
					return Result;
				}
			}
		}
	}
	return Result;
}
*/
__declspec(dllexport) double WINAPI PEM(FP DateRange[], FP TracerRange[], double Tau, double SampleDate, 
						double Lambda, double PEM_Uratio, double PEM_Lratio, double UZtime, LPXLOPER12 HeliumThree, LPXLOPER12 InitialTrit, LPXLOPER12 TritInitialTritRatio)
{
	double Result=0.0;
	double EndAge, BeginAge, nS, nU, nL, nStar;
	double TauRes, Tstar, pU, pS, pL, TauUpper, TauLower;
	double DR, Cin; // pointers to Date Range and Tracer Input
	double EndDate, PEM1=0.0, CinHe3, PEMnoDecay=0.0, TimeIncrement, MinDate, MaxDate;
	double PEMhalf1=0.0, PEMhalf2=0.0, Multiplier, MinAge, MaxAge, PEMnd1=0.0, PEMnd2=0.0, Lambda2;
	__int32 i,j, nIters, StepInc, StopCriteria;
//
	if (Tau > 100)
	{
		TimeIncrement = fabs(DateRange->array[1]-DateRange->array[0]); //MinTimeInc(DateRange);
		if (fabs(DateRange->array[DateRange->rows - 1]-DateRange->array[DateRange->rows - 2]) < TimeIncrement)
		{
			TimeIncrement = fabs(DateRange->array[DateRange->rows - 1]-DateRange->array[DateRange->rows - 2]);
		}
	}
	else
	{
		TimeIncrement = 1.0 / 12.0;
	}
	nIters = 2000000;
	nU = PEM_Uratio + 1.0;
	nL = PEM_Lratio + 1.0;
	if (PEM_Uratio == 0 && PEM_Lratio > 0) // PEM with screen at water table and above bottom of aquifer
	{
		nStar = 1.0 / PEM_Lratio + 1.0;
		nS = nStar;
		TauRes = Tau / (1 - (1 + log (nL)) / nL) / nStar;
		Tstar = log (nL) * TauRes;
		EndAge = Tstar + UZtime;
		BeginAge = UZtime;
		nIters = (__int32)((EndAge - BeginAge) / TimeIncrement) - 1;
	}
	else if (PEM_Lratio > PEM_Uratio && PEM_Lratio != 0) // PEM with screen below water table and above bottom of aquifer
	{
		nStar = 1 / PEM_Uratio + 1;
		pU = 1 / nStar;
		pL = 1 / nL;
		pS = 1 - pU - pL;
		nS = 1 / pS;
		TauUpper = (1 - log (nU) / nU - 1 / nU) * nStar * pU;
		TauLower = (log (nL) + 1) * pL;
		TauRes = pS * Tau / (1 - TauUpper - TauLower);
		Tstar = log (nU) * TauRes;
		BeginAge = Tstar + UZtime;
		Tstar = log (nL) * TauRes;
		EndAge = Tstar + UZtime;
		nIters = (__int32)((EndAge - BeginAge) / TimeIncrement) - 1;
	}
	else if (PEM_Uratio > 0 && PEM_Lratio == 0) // PEM with screen below water table and screened to bottom of aquifer
	{
		nS = nU;
		TauRes = Tau / (1 + log (nU));
		Tstar = log (nU) * TauRes;
		BeginAge = Tstar + UZtime;
		Tstar = log (nL) * TauRes;
		EndAge = 1e6;
	}
	else if (PEM_Uratio == 0 && PEM_Lratio == 0) // Full EMM
	{
		nS = 1;
		TauRes = Tau;
		BeginAge = UZtime;
		EndAge = 1e6;
	}
	else
	{
		return Result;
	}
	//
	MinDate = SampleDate - BeginAge; //Age at Z star
	MaxDate = floor (MinDate) + (floor ((MinDate - floor (MinDate)) / TimeIncrement)) * TimeIncrement;
	EndDate = MaxDate;
	MinAge = SampleDate - MinDate;
	MaxAge = SampleDate - MaxDate;
	//
	j = DateRange->rows - 1;
	StepInc = 1;
	StopCriteria = 0;
	if (DateRange->array[j] < DateRange->array[0])
	{
		StepInc = -1;
		StopCriteria = j;
		j = 0;
	}
	DR = DateRange->array[j];
	while ((DR >= EndDate) && j != StopCriteria)
	{
		j = j - StepInc;
		DR = DateRange->array[j];
	}
	Cin = TracerRange->array[j];
	if (Cin == 0 && j == StopCriteria)
	{
		return Result;
	}
	if (HeliumThree->val.xbool == TRUE)
	{
		if (Lambda == 0.0)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		CinHe3 = Cin * exp(-Lambda2 * UZtime) * (1 - exp(-Lambda2 * ((MaxAge-UZtime)-(MaxAge-MinAge)/2)));
		PEMhalf1 = exp(-(MinAge - UZtime) / TauRes);
		PEMhalf2 = exp(-(MaxAge - UZtime) / TauRes);
		Result = CinHe3 * nS * (PEMhalf1 - PEMhalf2);
		MinAge = MaxAge;
		for (i = 1; i<= nIters; i++)
		{		
			PEMhalf1 = PEMhalf2;
			MaxAge = MinAge + i * TimeIncrement;
			MaxDate = EndDate - i * TimeIncrement;
			while ((MaxDate < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					return Result;
				}
			}
			PEM1 = Result;
			CinHe3 = Cin * exp(-Lambda2 * UZtime) * (1 - exp(-Lambda2 * ((MaxAge - UZtime) - TimeIncrement / 2)));
			PEMhalf2 = exp(-(MaxAge - UZtime) / TauRes);
			Result += CinHe3 * nS * (PEMhalf1 - PEMhalf2);
			if (MaxAge >= Tau && Result > 0)
			{
				if ((Result - PEM1) / Result < Tol)
				{
					return Result;
				}
			}
		}
		i++;
		MaxAge = MinAge + i * TimeIncrement;
		if (MaxAge >= EndAge)
		{
			PEMhalf1 = PEMhalf2;
			MaxAge = EndAge;
			MaxDate = EndDate - i * TimeIncrement;
			while ((MaxDate < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					return Result;
				}
			}
			PEM1 = Result;
			CinHe3 = Cin * exp(-Lambda2 * UZtime) * (1 - exp(-Lambda2 * ((MaxAge - UZtime) - TimeIncrement / 2)));
			PEMhalf2 = exp(-(MaxAge - UZtime) / TauRes);
			Result += CinHe3 * nS * (PEMhalf1 - PEMhalf2);
		}
		return Result;
	}
	if (TritInitialTritRatio->val.xbool == TRUE)
	{
		if (Lambda == 0.0)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		Multiplier = nS / TauRes * (1 / ((1 / TauRes) + Lambda2));
		PEMnd1 = exp(-(MinAge - UZtime) / TauRes);
		PEMnd2 = exp(-(MaxAge - UZtime) / TauRes);
		PEMnoDecay = TracerRange->array[j] * exp(-Lambda2 * UZtime) * nS * (PEMnd1 - PEMnd2);
		PEMnd1 = PEMnd2;
		PEMhalf1 = exp(-(MinAge - UZtime) * ((1 / TauRes) + Lambda2));
		PEMhalf2 = exp(-(MaxAge - UZtime) * ((1 / TauRes) + Lambda2));
		Result = Cin * exp(-Lambda2 * UZtime) * Multiplier * (PEMhalf1 - PEMhalf2);
		MinAge = MaxAge;
		PEMhalf1 = PEMhalf2;
		for (i = 1; i<= nIters; i++)
		{
			PEM1 = Result;
			MaxAge = MinAge + i * TimeIncrement;
			MaxDate = EndDate - i * TimeIncrement;
			while ((MaxDate < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					break;
				}
			}
			PEMnd2 = exp(-(MaxAge - UZtime) / TauRes);
			PEMnoDecay += Cin * exp(-Lambda2 * UZtime) * nS * (PEMnd1 - PEMnd2);
			PEMnd1 = PEMnd2;
			PEMhalf2 = exp(-(MaxAge - UZtime) * ((1 / TauRes) + Lambda2));
			Result += Cin * exp(-Lambda2 * UZtime) * Multiplier * (PEMhalf1 - PEMhalf2);
			PEMhalf1 = PEMhalf2;
			if (MaxAge >= Tau && Result > 0)
			{
				if ((Result - PEM1) / Result < Tol)
				{
					break;
				}
			}
		}
		i++;
		MaxAge = MinAge + i * TimeIncrement;
		if (MaxAge >= EndAge)
		{
			PEM1 = Result;
			MaxAge = EndAge;
			MaxDate = EndDate - i * TimeIncrement;
			while ((MaxDate < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					return Result;
				}
			}
			PEMnd2 = exp(-(MaxAge - UZtime) / TauRes);
			PEMnoDecay += Cin * exp(-Lambda2 * UZtime) * nS * (PEMnd1 - PEMnd2);
			PEMnd1 = PEMnd2;
			PEMhalf2 = exp(-(MaxAge - UZtime) * ((1 / TauRes) + Lambda2));
			Result += Cin * exp(-Lambda2 * UZtime) * Multiplier * (PEMhalf1 - PEMhalf2);
			PEMhalf1 = PEMhalf2;
		}
		Result = Result / PEMnoDecay;
		return Result;
	}
	if (TritInitialTritRatio->val.xbool != TRUE && HeliumThree->val.xbool != TRUE)
	{
		if (InitialTrit->val.xbool == TRUE)
		{
			Lambda2 = (double) log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		Multiplier = nS / TauRes * (1 / ((1 / TauRes) + Lambda));
		PEMhalf1 = exp(-(MinAge - UZtime) * ((1 / TauRes) + Lambda));
		PEMhalf2 = exp(-(MaxAge - UZtime) * ((1 / TauRes) + Lambda));
		Result = Cin * exp(-Lambda2 * UZtime) * Multiplier * (PEMhalf1 - PEMhalf2);
		MinAge = MaxAge;
		for (i = 1; i<= nIters; i++)
		{
			PEMhalf1 = PEMhalf2;
			MaxAge = MinAge + i * TimeIncrement;
			MaxDate = EndDate - i * TimeIncrement;
			while ((MaxDate < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					return Result;
				}
			}
			PEM1 = Result;
			PEMhalf2 = exp(-(MaxAge - UZtime) * ((1 / TauRes) + Lambda));
			Result += Cin * exp(-Lambda2 * UZtime) * Multiplier * (PEMhalf1 - PEMhalf2);
			if (MaxAge >= Tau && Result > 0)
			{
				if ((Result - PEM1) / Result < Tol)
				{
					return Result;
				}
			}
		}
		i++;
		MaxAge = MinAge + i * TimeIncrement;
		if (MaxAge >= EndAge)
		{
			PEMhalf1 = PEMhalf2;
			MaxAge = EndAge;
			MaxDate = EndDate - i * TimeIncrement;
			while ((MaxDate < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					return Result;
				}
			}
			PEM1 = Result;
			PEMhalf2 = exp(-(MaxAge - UZtime) * ((1 / TauRes) + Lambda));
			Result += Cin * exp(-Lambda2 * UZtime) * Multiplier * (PEMhalf1 - PEMhalf2);
		}
		return Result;
	}
	return Result;
}
/*__declspec(dllexport) double WINAPI PEM_He4(
                        double Uppm, double THppm, double Porosity, double SedDensity, double Tau, double PEMratio, double HeSolnRate)
{
	double Result;
	double Cin, PEMold, TimeIncrement, MinAge, MaxAge, n, TauRes, PEMhalf1, PEMhalf2;
	__int32 i, SimYears, nIters;
    
    n = PEMratio + 1;
    TauRes = Tau / (1 + log (n));
    SimYears = 1500000;
    if (Tau < 100)
	{
		TimeIncrement = 1.0 / 12.0;
	}
	else
	{
		if (Tau < 1000)
		{
			TimeIncrement = 0.5;
		}
		else
		{
			TimeIncrement = Tau / 1000;
		}
	}
    MinAge = log (n) * TauRes; //Tstar
    MaxAge = floor (MinAge) + (floor ((MinAge-floor (MinAge)) / TimeIncrement) + 1) * TimeIncrement;
    nIters = (__int32) ((SimYears - MinAge) / TimeIncrement);
    Result = 0.0;
	if (HeSolnRate != 0.0)
	{
		Cin = HeSolnRate;
	}
	else
	{
		Cin = SedDensity / Porosity * (Udecay * Uppm + THdecay * THppm);
	}
    PEMhalf1 = exp(-MinAge / TauRes);
    PEMhalf2 = exp(-MaxAge / TauRes);
    Result = Cin * MaxAge * n * (PEMhalf1 - PEMhalf2);
	MinAge = MaxAge;
    for (i = 1; i<= nIters; i++)
	{
        PEMold = Result;
		PEMhalf1 = PEMhalf2;
        MaxAge = i * TimeIncrement + MinAge;
        PEMhalf2 = exp(-MaxAge / TauRes);
        Result += Cin * MaxAge * n * (PEMhalf1 - PEMhalf2);
        if (MaxAge >= Tau && Result > 0)
			{
				if ((Result - PEMold) / Result < Tol)
				{
					return Result;
				}
			}
	}
	return Result;
}
*/
__declspec(dllexport) double WINAPI PEM_He4(
                        double Uppm, double THppm, double Porosity, double SedDensity, double Tau, double PEM_Uratio, double PEM_Lratio, double HeSolnRate)
{
	double Result=0.0;
	double EndAge, BeginAge, nS, nU, nL, nStar;
	double TauRes, Tstar, pU, pS, pL, TauUpper, TauLower;
	double Cin, TimeIncrement, n;
	double PEMhalf1=0.0, PEMhalf2=0.0, PEMold, Multiplier, MinAge, MaxAge;
	__int32 i, nIters, SimYears;
    
    nIters = 2000000;
	SimYears = 1500000;
    if (Tau < 100)
	{
		TimeIncrement = 1.0 / 12.0;
	}
	else
	{
		if (Tau < 1000)
		{
			TimeIncrement = 0.5;
		}
		else
		{
			TimeIncrement = Tau / 1000;
		}
	}
	nU = PEM_Uratio + 1.0;
	nL = PEM_Lratio + 1.0;
	if (PEM_Uratio == 0 && PEM_Lratio > 0) // PEM with screen at water table and above bottom of aquifer
	{
		nStar = 1.0 / PEM_Lratio + 1.0;
		nS = nStar;
		TauRes = Tau / (1 - (1 + log (nL)) / nL) / nStar;
		Tstar = log (nL) * TauRes;
		EndAge = Tstar;
		BeginAge = 0.0;
		nIters = (__int32)((EndAge - BeginAge) / TimeIncrement) - 1;
	}
	else if (PEM_Lratio > PEM_Uratio && PEM_Lratio != 0) // PEM with screen below water table and above bottom of aquifer
	{
		nStar = 1 / PEM_Uratio + 1;
		pU = 1 / nStar;
		pL = 1 / nL;
		pS = 1 - pU - pL;
		nS = 1 / pS;
		TauUpper = (1 - log (nU) / nU - 1 / nU) * nStar * pU;
		TauLower = (log (nL) + 1) * pL;
		TauRes = pS * Tau / (1 - TauUpper - TauLower);
		Tstar = log (nU) * TauRes;
		BeginAge = Tstar;
		Tstar = log (nL) * TauRes;
		EndAge = Tstar;
		nIters = (__int32)((EndAge - BeginAge) / TimeIncrement) - 1;
	}
	else if (PEM_Uratio > 0 && PEM_Lratio == 0) // PEM with screen below water table and screened to bottom of aquifer
	{
		nS = nU;
		TauRes = Tau / (1 + log (nU));
		Tstar = log (nU) * TauRes;
		BeginAge = Tstar;
		Tstar = log (nL) * TauRes;
		EndAge = 1e6;
	}
	else if (PEM_Uratio == 0 && PEM_Lratio == 0) // Full EMM
	{
		nS = 1;
		TauRes = Tau;
		BeginAge = 0.0;
		EndAge = 1e6;
	}
	else
	{
		return Result;
	}
	//
	MinAge = BeginAge; //Tstar
    MaxAge = floor (MinAge) + (floor ((MinAge-floor (MinAge)) / TimeIncrement) + 1) * TimeIncrement;
	if (HeSolnRate != 0.0)
	{
		Cin = HeSolnRate;
	}
	else
	{
		Cin = SedDensity / Porosity * (Udecay * Uppm + THdecay * THppm);
	}
	PEMhalf1 = exp(-MinAge / TauRes);
	PEMhalf2 = exp(-MaxAge / TauRes);
	Result = Cin * MaxAge * nS * (PEMhalf1 - PEMhalf2);
	MinAge = MaxAge;
	for (i = 1; i<= nIters; i++)
	{
		PEMhalf1 = PEMhalf2;
		MaxAge = MinAge + i * TimeIncrement;
		PEMold = Result;
		PEMhalf2 = exp(-MaxAge / TauRes);
		Result += Cin * nS * MaxAge * (PEMhalf1 - PEMhalf2);
		if (MaxAge >= Tau && Result > 0)
		{
			if ((Result - PEMold) / Result < Tol)
			{
				return Result;
			}
		}
	}
	i++;
	MaxAge = MinAge + i * TimeIncrement;
	if (MaxAge >= EndAge)
	{
		PEMhalf1 = PEMhalf2;
		MaxAge = EndAge;
		PEMold = Result;
		PEMhalf2 = exp(-MaxAge / TauRes);
		Result += Cin * nS * (PEMhalf1 - PEMhalf2);
	}
	return Result;
    /*PEMhalf1 = exp(-MinAge / TauRes);
    PEMhalf2 = exp(-MaxAge / TauRes);
    Result = Cin * MaxAge * nS * (PEMhalf1 - PEMhalf2);
	MinAge = MaxAge;
    for (i = 1; i<= nIters; i++)
	{
        PEMold = Result;
		PEMhalf1 = PEMhalf2;
        MaxAge = i * TimeIncrement + MinAge;
        PEMhalf2 = exp(-MaxAge / TauRes);
        Result += Cin * MaxAge * nS * (PEMhalf1 - PEMhalf2);
        if (MaxAge >= Tau && Result > 0)
			{
				if ((Result - PEMold) / Result < Tol)
				{
					return Result;
				}
			}
	}
	return Result;*/
}

__declspec(dllexport) double WINAPI EPM(
                        FP DateRange[], FP TracerRange[], double Tau, double SampleDate, 
						double Lambda, double EPMratio, double UZtime, LPXLOPER12 HeliumThree, LPXLOPER12 InitialTrit, LPXLOPER12 TritInitialTritRatio)
{
	double Result;
	double DR, Cin; // pointers to Date Range and Tracer Input
	double n, EndDate, EPM1, CinHe3, EPMnoDecay, TimeIncrement;
	double EPMhalf1, EPMhalf2, Multiplier, MinAge, MaxAge, EPMnd1, EPMnd2, Lambda2;
	//bool TorF; // pointer to boolean values
	__int32 i,j, nIters, StepInc, StopCriteria;

	Result = 0;
	n = EPMratio + 1;
	if (Tau > 100.)
	{
        TimeIncrement = fabs(DateRange->array[1]-DateRange->array[0]); //MinTimeInc(DateRange);
		if (fabs(DateRange->array[DateRange->rows - 1]-DateRange->array[DateRange->rows - 2]) < TimeIncrement)
		{
			TimeIncrement = fabs(DateRange->array[DateRange->rows - 1]-DateRange->array[DateRange->rows - 2]);
		}
	}
    else
	{
        TimeIncrement = 1.0 / 12.0;
	}
	MinAge = SampleDate - Tau*(1.0 - (1.0 / n));
	MaxAge = floor (MinAge) + (floor ((MinAge - floor (MinAge)) / TimeIncrement)) * TimeIncrement;
	EndDate = SampleDate - UZtime;
	MinAge = SampleDate - MinAge;
	MaxAge = SampleDate - MaxAge;
	nIters = 2000000;
	j = DateRange->rows - 1;
	StepInc = 1;
	StopCriteria = 0;
	if (DateRange->array[j] < DateRange->array[0])
	{
		StepInc = -1;
		StopCriteria = j;
		j = 0;
	}
	DR = DateRange->array[j];
	while ((DR >= EndDate - MinAge) && j != StopCriteria)
	{
		j = j - StepInc;
		DR = DateRange->array[j];
	}
	Cin = TracerRange->array[j];
	if (Cin == 0 && j == StopCriteria)
	{
		return Result;
	}
	//TorF = HeliumThree->val.num
	if (HeliumThree->val.xbool == TRUE)
	{
		if (Lambda == 0.0)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		CinHe3 = Cin * exp(-Lambda2 * UZtime) * (1 - exp(-Lambda2 * (MaxAge - (MaxAge-MinAge) / 2)));
		EPMhalf1 = exp(-MinAge * n / Tau + n - 1.0);
		EPMhalf2 = exp(-MaxAge * n / Tau + n - 1.0);
		Result = CinHe3 * (EPMhalf1 - EPMhalf2);
		MinAge = MaxAge;
		if (Cin == 0 && j == StopCriteria)
		{
			return Result;
		}
		for (i = 1; i<= nIters; i++)
		{		
			EPMhalf1 = EPMhalf2;
			MaxAge = MinAge + i * TimeIncrement;
			while ((EndDate - MaxAge < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					return Result;
				}
			}
			EPM1 = Result;
			CinHe3 = Cin * exp(-Lambda2 * UZtime) * (1 - exp(-Lambda2 * (MaxAge - TimeIncrement / 2)));
			EPMhalf2 = exp(-MaxAge * n / Tau + n - 1.0);
			Result += CinHe3 * (EPMhalf1 - EPMhalf2);
			if (MaxAge >= Tau && Result > 0)
			{
				if ((Result - EPM1) / Result < Tol)
				{
					return Result;
				}
			}
		}
	}
	if (TritInitialTritRatio->val.xbool == TRUE)
	{
		if (Lambda == 0.0)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		Multiplier = n / Tau * (1.0 / ((n / Tau) + Lambda2));
		EPMnd1 = exp(-MinAge * n / Tau + n - 1.0);
		EPMnd2 = exp(-MaxAge * n / Tau + n - 1.0);
		EPMnoDecay = TracerRange->array[j] * exp(-Lambda2 * UZtime) * (EPMnd1 - EPMnd2);
		EPMnd1 = EPMnd2;
		EPMhalf1 = exp(-MinAge * (n / Tau + Lambda2) + n - 1.0);
		EPMhalf2 = exp(-MaxAge* (n / Tau + Lambda2) + n - 1.0);
		Result = Cin * exp(-Lambda2 * UZtime) * Multiplier * (EPMhalf1 - EPMhalf2);
		EPMhalf1 = EPMhalf2;
		MinAge = MaxAge;
		for (i = 1; i<= nIters; i++)
		{
			EPM1 = Result;
			MaxAge = MinAge + i * TimeIncrement;
			while ((EndDate - MaxAge < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					break;
				}
			}
			EPMnd2 = exp(-MaxAge * n / Tau + n - 1.0);
			EPMnoDecay += + Cin * exp(-Lambda2 * UZtime) * (EPMnd1 - EPMnd2);
			EPMnd1 = EPMnd2;
			EPMhalf2 = exp(-MaxAge * (n / Tau + Lambda2) + n - 1.0);
			Result += Cin * exp(-Lambda2 * UZtime) * Multiplier * (EPMhalf1 - EPMhalf2);
			EPMhalf1 = EPMhalf2;
			if (MaxAge >= Tau && Result > 0)
			{
				if ((Result - EPM1) / Result < Tol)
				{
					break;
				}
			}
		}
		Result = Result / EPMnoDecay;
		return Result;
	}
	if (TritInitialTritRatio->val.xbool != TRUE && HeliumThree->val.xbool != TRUE)
	{
		if (InitialTrit->val.xbool == TRUE)
		{
			Lambda2 = log (2.0) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		Multiplier = n / Tau * (1 / ((n / Tau) + Lambda));
		EPMhalf1 = exp(-MinAge * (n / Tau + Lambda) + n - 1.0);
		EPMhalf2 = exp(-MaxAge * (n / Tau + Lambda) + n - 1.0);
		Result = Cin * exp(-Lambda2 * UZtime) * Multiplier * (EPMhalf1 - EPMhalf2);
		MinAge = MaxAge;
		for (i = 1; i<= nIters; i++)
		{
			EPMhalf1 = EPMhalf2;
			MaxAge = MinAge + i * TimeIncrement;
			while ((EndDate - MaxAge < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					return Result;
				}
			}
			EPM1 = Result;
			EPMhalf2 = exp(-MaxAge * (n / Tau + Lambda) + n - 1.0);
			Result += Cin * exp(-Lambda2 * UZtime) * Multiplier * (EPMhalf1 - EPMhalf2);
			if (MaxAge >= Tau && Result > 0)
			{
				if ((Result - EPM1) / Result < Tol)
				{
					return Result;
				}
			}
		}
	}
	return Result;
}

__declspec(dllexport) double WINAPI EPM_He4(
                        double Uppm, double THppm, double Porosity, double SedDensity,double Tau, double EPMratio, double HeSolnRate)
{
	double Result;
	double Cin, EPMold, TimeIncrement, MinAge, MaxAge, n, EPMhalf1, EPMhalf2;
	__int32 i, SimYears, nIters;
    
    n = EPMratio + 1;
    SimYears = 1500000;
    if (Tau < 100)
	{
		TimeIncrement = 1.0 / 12.0;
	}
	else
	{
		if (Tau < 1000)
		{
			TimeIncrement = 0.5;
		}
		else
		{
			TimeIncrement = Tau / 1000;
		}
	}
    MinAge = Tau * (1.0 - (1.0 / n));
    MaxAge = floor (MinAge) + (floor ((MinAge-floor (MinAge)) / TimeIncrement) + 1) * TimeIncrement;
    nIters = (__int32) ((SimYears - MinAge) / TimeIncrement);
    Result = 0.0;
	if (HeSolnRate != 0.0)
	{
		Cin = HeSolnRate;
	}
	else
	{
		Cin = SedDensity / Porosity * (Udecay * Uppm + THdecay * THppm);
	}
    EPMhalf1 = exp(-MinAge * n / Tau + n - 1.0);
    EPMhalf2 = exp(-MaxAge * n / Tau + n - 1.0);
    Result = Cin * MaxAge * (EPMhalf1 - EPMhalf2);
	MinAge = MaxAge;
    for (i = 1; i<= nIters; i++)
	{
        EPMold = Result;
		EPMhalf1 = EPMhalf2;
        MaxAge = i * TimeIncrement + MinAge;
        EPMhalf2 = exp(-MaxAge * n / Tau + n - 1.0);
        Result += Cin * MaxAge * (EPMhalf1 - EPMhalf2);
        if (MaxAge >= Tau && Result > 0)
			{
				if ((Result - EPMold) / Result < Tol)
				{
					return Result;
				}
			}
	}
	return Result;
}

double ReturnExcelGamma(double age, double alpha, double beta)
{
	XLOPER12 Result, Beta, GamBool, GamAge, Alpha;

	Result.xltype=xltypeNum|xlbitDLLFree;
	Beta.xltype=xltypeNum|xlbitDLLFree;
	GamBool.xltype=xltypeNum|xlbitDLLFree;
	GamAge.xltype=xltypeNum|xlbitDLLFree;
	Alpha.xltype=xltypeNum|xlbitDLLFree;

	Result.val.num = 0.0;
	Beta.val.num=beta;
	Alpha.val.num=alpha;
	GamAge.val.num=age;
	GamBool.val.num=1;

	Excel12f(495,&Result,4,&GamAge,&Alpha,&Beta,&GamBool);

	return Result.val.num;
}

__declspec(dllexport) double WINAPI GAM(
                        FP DateRange[], FP TracerRange[], double Tau, double SampleDate, 
						double Lambda, double Alpha, double UZtime, bool HeliumThree, bool InitialTrit, bool TritInitialTritRatio)
{
	
	double DR, Cin, Beta; // pointers to Date Range and Tracer Input
	double EndDate, GAM1, CinHe3, GAMnoDecay, TimeIncrement, UZdecay, MaxDecay;
	double GAMhalf1, GAMhalf2, MinAge, MaxAge, GAMnd1, GAMnd2, Lambda2, RtnRslt=0.0;
	//bool TorF; // pointer to boolean values
	__int32 i,j, nIters, StepInc, StopCriteria;

	
	
	Beta = Tau/Alpha;
	if (Tau > 100.)
	{
        TimeIncrement = fabs(DateRange->array[1]-DateRange->array[0]); //MinTimeInc(DateRange);
		if (fabs(DateRange->array[DateRange->rows - 1]-DateRange->array[DateRange->rows - 2]) < TimeIncrement)
		{
			TimeIncrement = fabs(DateRange->array[DateRange->rows - 1]-DateRange->array[DateRange->rows - 2]);
		}
	}
    else
	{
        TimeIncrement = 1.0 / 12.0;
	}
	MinAge = SampleDate;
	MaxAge = floor (MinAge) + (floor ((MinAge - floor (MinAge)) / TimeIncrement)) * TimeIncrement;
	EndDate = SampleDate - UZtime;
	MinAge = SampleDate - MinAge;
	MaxAge = SampleDate - MaxAge;
	nIters = 2000000;
	j = DateRange->rows - 1;
	StepInc = 1;
	StopCriteria = 0;
	if (DateRange->array[j] < DateRange->array[0])
	{
		StepInc = -1;
		StopCriteria = j;
		j = 0;
	}
	DR = DateRange->array[j];
	while ((DR >= EndDate - MinAge) && j != StopCriteria)
	{
		j = j - StepInc;
		DR = DateRange->array[j];
	}
	Cin = TracerRange->array[j];
	if (Cin == 0 && j == StopCriteria)
	{
		return RtnRslt;
	}
	//TorF = HeliumThree->val.num
	if (HeliumThree == TRUE)
	{
		if (Lambda == 0.0)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		MaxDecay = MaxAge-TimeIncrement/2;
		UZdecay = exp(-Lambda2 * UZtime);
		CinHe3 = Cin * UZdecay * (1 - exp(-Lambda2 * MaxDecay));
	
		GAMhalf1=ReturnExcelGamma(MinAge,Alpha,Beta);
		GAMhalf2=ReturnExcelGamma(MaxAge,Alpha,Beta);

		RtnRslt = CinHe3 * (GAMhalf2 - GAMhalf1);
		MinAge = MaxAge;
		if (Cin == 0 && j == StopCriteria)
		{
			return RtnRslt;
		}
		for (i = 1; i<= nIters; i++)
		{		
			GAMhalf1 = GAMhalf2;
			MaxAge = MinAge + i * TimeIncrement;
			MaxDecay = MaxAge-TimeIncrement/2;
			while ((EndDate - MaxAge < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					return RtnRslt;
				}
			}
			GAM1 = RtnRslt;
			CinHe3 = Cin * UZdecay * (1 - exp(-Lambda2 * MaxDecay));
			
			
			GAMhalf2 = ReturnExcelGamma(MaxAge,Alpha,Beta);
			RtnRslt += CinHe3 * (GAMhalf2 - GAMhalf1);
			if (MaxAge >= Tau && RtnRslt > 0)
			{
				if ((RtnRslt - GAM1) / RtnRslt < Tol)
				{
					return RtnRslt;
				}
			}
		}
	}
	if (TritInitialTritRatio == TRUE)
	{
		if (Lambda == 0.0)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
	
		MaxDecay = MaxAge-TimeIncrement/2;
		UZdecay = exp(-Lambda2 * UZtime);
		
		GAMnd1 = ReturnExcelGamma(MinAge,Alpha,Beta);
		
		
		GAMnd2 = ReturnExcelGamma(MaxAge,Alpha,Beta);
		GAMnoDecay = TracerRange->array[j] * exp(-Lambda2 * UZtime) * (GAMnd2 - GAMnd1);
		GAMnd1 = GAMnd2;
		
		GAMhalf1 = ReturnExcelGamma(MinAge,Alpha,Beta);
		
		GAMhalf2 = ReturnExcelGamma(MaxAge,Alpha,Beta);
		RtnRslt = Cin * UZdecay * (GAMhalf2 - GAMhalf1)*exp(-Lambda2 * MaxDecay);
		GAMhalf1 = GAMhalf2;
		MinAge = MaxAge;
		for (i = 1; i<= nIters; i++)
		{
			GAM1 = RtnRslt;
			MaxAge = MinAge + i * TimeIncrement;
			MaxDecay = MaxAge-TimeIncrement/2;
			while ((EndDate - MaxAge < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					break;
				}
			}
			
			GAMnd2 = ReturnExcelGamma(MaxAge,Alpha,Beta);
			GAMnoDecay +=  Cin * UZdecay * (GAMnd2 - GAMnd1);
			GAMnd1 = GAMnd2;
			
			GAMhalf2 = ReturnExcelGamma(MinAge,Alpha,Beta);
			RtnRslt += Cin * UZdecay *(GAMhalf2 - GAMhalf1)*exp(-Lambda2 * MaxDecay);
			GAMhalf1 = GAMhalf2;
			if (MaxAge >= Tau && RtnRslt > 0)
			{
				if ((RtnRslt - GAM1) / RtnRslt < Tol)
				{
					break;
				}
			}
		}
		RtnRslt = RtnRslt / GAMnoDecay;
		return RtnRslt;
	}
	if (TritInitialTritRatio != TRUE && HeliumThree != TRUE)
	{
		if (InitialTrit == TRUE)
		{
			Lambda2 = log (2.0) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		
		
		UZdecay = exp(-Lambda2 * UZtime);
		
		GAMhalf1 = ReturnExcelGamma(MinAge,Alpha,Beta);
		
		GAMhalf2 = ReturnExcelGamma(MaxAge,Alpha,Beta);
		MaxDecay = MaxAge-TimeIncrement/2;
		RtnRslt = Cin * UZdecay *(GAMhalf2 - GAMhalf1)*exp(-Lambda * MaxDecay);
		MinAge = MaxAge;
		for (i = 1; i<= nIters; i++)
		{
			GAMhalf1 = GAMhalf2;
			MaxAge = MinAge + i * TimeIncrement;
			MaxDecay = MaxAge - TimeIncrement / 2;
			while ((EndDate - MaxAge < DR) && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					return RtnRslt;
				}
			}
			GAM1 = RtnRslt;
			
			GAMhalf2 = ReturnExcelGamma(MaxAge,Alpha,Beta);
			RtnRslt += Cin * UZdecay * (GAMhalf2 - GAMhalf1)*exp(-Lambda * MaxDecay);
			if (MaxAge >= Tau && RtnRslt > 0)
			{
				if ((RtnRslt - GAM1) / RtnRslt < Tol)
				{
					return RtnRslt;
				}
			}
		}
	}
	return RtnRslt;
}

__declspec(dllexport) double WINAPI GAM_He4(
                        double Uppm, double THppm, double Porosity, double SedDensity,double Tau, double Alpha,  double HeSolnRate)
{
	double Result, Beta;
	double Cin, GAMold, TimeIncrement, MinAge, MaxAge, GAMhalf1, GAMhalf2;
	__int32 i, SimYears, nIters;

    Beta = Tau / Alpha;
    SimYears = 1500000;
    if (Tau < 100)
	{
		TimeIncrement = 1.0 / 12.0;
	}
	else
	{
		if (Tau < 1000)
		{
			TimeIncrement = 0.5;
		}
		else
		{
			TimeIncrement = Tau / 1000;
		}
	}
    MinAge = 0.0;
    MaxAge = floor (MinAge) + (floor ((MinAge-floor (MinAge)) / TimeIncrement) + 1) * TimeIncrement;
    nIters = (__int32) ((SimYears - MinAge) / TimeIncrement);
    Result = 0.0;
	if (HeSolnRate != 0.0)
	{
		Cin = HeSolnRate;
	}
	else
	{
		Cin = SedDensity / Porosity * (Udecay * Uppm + THdecay * THppm);
	}
	GAMhalf1 = ReturnExcelGamma(MinAge,Alpha,Beta);
    GAMhalf2 = ReturnExcelGamma(MaxAge,Alpha,Beta);
    Result = Cin * MaxAge * (GAMhalf2-GAMhalf1);
	MinAge = MaxAge;
    for (i = 1; i<= nIters; i++)
	{
        MaxAge = i * TimeIncrement + MinAge;
		GAMhalf1 = GAMhalf2;
		GAMhalf2 = ReturnExcelGamma(MaxAge,Alpha,Beta);
        GAMold=Result; 
		Result += Cin * MaxAge * (GAMhalf2-GAMhalf1);
        if (MaxAge >= Tau && Result > 0)
			{
				if ((Result - GAMold) / Result < Tol)
				{
					return Result;
				}
			}
	}
	return Result;
}


//static const double rel_error= 1E-12;        //calculate 12 significant figures
////you can adjust rel_error to trade off between accuracy and speed
////but don't ask for > 15 figures (assuming usual 52 bit mantissa in a double)

//double erfc(double x)
////erfc(x) = 2/sqrt(pi)*integral(exp(-t^2),t,x,inf)
////        = exp(-x^2)/sqrt(pi) * [1/x+ (1/2)/x+ (2/2)/x+ (3/2)/x+ (4/2)/x+ ...]
////        = 1-erf(x)
////expression inside [] is a continued fraction so '+' means add to denominator
////only
//{
//	static const double one_sqrtpi=  0.564189583547756287;        // 1/sqrt(pi)
//	double a=1.0, b=x;                //last two convergent numerators
//	double c=x, d=x*x+0.5;          //last two convergent denominators
//	double q1, q2= b/d;             //last two convergents (a/c and b/d)
//	double n= 1.0, t;
//	if (fabs(x) < 2.2) 
//	{
//		return 1.0 - erf2(x);        //use series when fabs(x) < 2.2
//	}
//	if (x < 0.0) 
//	{               //continued fraction only valid for x>0
//		return 2.0 - erfc(-x);
//	}
//	do 
//	{
//		t= a*n+b*x;
//		a= b;
//		b= t;
//		t= c*n+d*x;
//		c= d;
//		d= t;
//		n+= 0.5;
//		q1= q2;
//		q2= b/d;
//	} while (fabs(q1-q2)/q2 > rel_error);
//	return one_sqrtpi*exp(-x*x)*q2;
//}

//double erf(double x)
////erf(x) = 2/sqrt(pi)*integral(exp(-t^2),t,0,x)
////       = 2/sqrt(pi)*[x - x^3/3 + x^5/5*2! - x^7/7*3! + ...]
////       = 1-erfc(x)
//{
//	static const double two_sqrtpi=  1.128379167095512574;        // 2/sqrt(pi)
//	double sum = x;
//	double term = x;
//	double xsqr = x*x;
//	int j= 1;
//	if (fabs(x) > 2.2) 
//	{
//		return 1.0 - erfc(x);        //use continued fraction when fabs(x) > 2.2
//	}
//	do 
//	{
//		term*= xsqr/j;
//		sum-= term/(2*j+1);
//		++j;
//		term*= xsqr/j;
//		sum+= term/(2*j+1);
//		++j;
//	} while (fabs(term/sum) > rel_error);   // CORRECTED LINE
//	return two_sqrtpi*sum;
//}

//double DMfunc(double a, double b, double x)
//{
//	double Result2 = 0.0;
//	Result2 = exp(2 * sqrt (a * b)) * erf(sqrt (a) * x + sqrt (b) / x) - exp(-2 * sqrt (a * b)) * erf(sqrt (a) * x - sqrt (b) / x);
//	return Result2;
//}

//double gt_DMold(double MinAge, double MaxAge, double Tau, double Lambda, double DP, double UZtime)
//{
//	double Result = 0.0, a, b, x1, x2, Dum1, Dum2;
//	if (MinAge > UZtime)
//	{
//		b = 1/ (4 * DP);
//		a = Lambda * Tau + b;
//		x1 = sqrt ((MaxAge - UZtime) / Tau);
//		x2 = sqrt ((MinAge - UZtime) / Tau);
//		Dum1 = DMfunc(a, b, x1);
//		Dum2 = DMfunc(a, b, x2);
//		Result = exp(1 / (2 * DP)) / (4 * sqrt (b * DP)) * (Dum2 - Dum1);
//	}
//	return Result;
//}

//double gt_DMintOld(double MinAge, double MaxAge, double Tau, double Lambda, double DP)
//{
//	double Result = 0.0, a, b, x1, x2, Dum1, Dum2;
//	b = 1/ (4 * DP);
//	a = Lambda * Tau + b;
//	x1 = sqrt (MaxAge / Tau);
//	x2 = sqrt (MinAge / Tau);
//	Dum1 = DMfunc(a, b, x1);
//	Dum2 = DMfunc(a, b, x2);
//	Result = exp(1 / (2 * DP)) / (4 * sqrt (b * DP)) * (Dum2 - Dum1);
//	return Result;
//}
double DispInt(double X, double Tau, double Lambda, double DP)
{
	double alpha, beta, phi, theta, Result3 = 0.0;
	phi = 4.0 * DP * PI;
	theta = pow ((X / Tau),3.0);
	alpha = 1.0 / (Tau * sqrt (phi) * sqrt (theta));
    beta = -(Lambda * Tau + 1.0 / (4.0 * DP)) * X / Tau - 1.0 / (4.0 * DP) * Tau / X + 1.0 / (2.0 * DP);
    Result3 = alpha * exp(beta);
    return Result3;
}
double DispInt2(double X, double Tau, double Lambda, double DP)
{
	double alpha, beta, phi, theta, Result3 = 0.0;
	phi = 4.0 * DP * PI;
	theta = pow ((X / Tau),3.0);
	alpha = 1.0 / (Tau * sqrt (phi) * sqrt (theta));
    beta = -(Lambda * Tau + 1.0 / (4.0 * DP)) * X / Tau - 1.0 / (4.0 * DP) * Tau / X + 1.0 / (2.0 * DP);
    Result3 = alpha * exp(beta);
    return Result3;
}
double DispInt3(double X, double Tau, double Lambda, double DP)
{
	double alpha, beta, phi, theta, Result3 = 0.0;
	phi = 4.0 * DP * PI;
	theta = pow ((X / Tau),3.0);
	alpha = 1.0 / (Tau * sqrt (phi) * sqrt (theta));
    beta = -(Lambda * Tau + 1.0 / (4.0 * DP)) * X / Tau - 1.0 / (4.0 * DP) * Tau / X + 1.0 / (2.0 * DP);
    Result3 = alpha * exp(beta);
    return Result3;
}
double gt_DMint(double MinAge, double MaxAge, double Tau, double Lambda, double DP)
{
    double Result=0.0;
	if ((MaxAge == 0) | (DP == 0))
	{
		return Result;
	}
	//double a=MinAge-UZtime, b=MaxAge-UZtime;
	double k1 = (MaxAge - MinAge)/2, k2 = (MinAge + MaxAge)/2;
	Result+=k1*Coeff5[0]*DispInt2(k1*Roots5[0]+k2,Tau,Lambda,DP);
	Result+=k1*Coeff5[1]*DispInt2(k1*Roots5[1]+k2,Tau,Lambda,DP);
	Result+=k1*Coeff5[2]*DispInt2(k1*Roots5[2]+k2,Tau,Lambda,DP);
	Result+=k1*Coeff5[3]*DispInt2(k1*Roots5[3]+k2,Tau,Lambda,DP);
	Result+=k1*Coeff5[4]*DispInt2(k1*Roots5[4]+k2,Tau,Lambda,DP);
	return Result;
}
double gt_DMint2(double MinAge, double MaxAge, double Tau, double Lambda, double DP)
{
    int i=1;
    double Aarray[50], TolArray[50], h[50], FA[50], FB[50], FC[50], S[50], L[50];
    double FD, FE, S1, S2, v1, v2, v3, v4, v5, v6, v7, v8, Result2 = 0.0, Tolerance=1e-4;
    TolArray[i] = 10 * Tolerance;
    Aarray[i] = MinAge;
    h[i] = (MaxAge - MinAge) / 2;
    FA[i] = DispInt3(MinAge, Tau, Lambda, DP);
    FC[i] = DispInt3(MinAge + h[i], Tau, Lambda, DP);
    FB[i] = DispInt3(MaxAge, Tau, Lambda, DP);
    S[i] = h[i] * (FA[i] + 4 * FC[i] + FB[i]) / 3;
    L[i] = 1;
    do
	{
        FD = DispInt3(Aarray[i] + h[i] / 2, Tau, Lambda, DP);
        FE = DispInt3(Aarray[i] + 3 * h[i] / 2, Tau, Lambda, DP);
        S1 = h[i] * (FA[i] + 4 * FD + FC[i]) / 6;
        S2 = h[i] * (FC[i] + 4 * FE + FB[i]) / 6;
        v1 = Aarray[i];
        v2 = FA[i];
        v3 = FC[i];
        v4 = FB[i];
        v5 = h[i];
        v6 = TolArray[i];
        v7 = S[i];
        v8 = L[i];
        i--;
        if (fabs(S1 + S2 - v7) < v6)
		{
            Result2 += (S1 + S2);
			//return Result2;
		}
        else if (i <= 47)
		{
            i++;
            Aarray[i] = v1 + v5;
            FA[i] = v3;
            FC[i] = FE;
            FB[i] = v4;
            h[i] = v5 / 2;
            TolArray[i] = v6 / 2;
            S[i] = S2;
            L[i] = v8 + 1;
            i++;
            Aarray[i] = v1;
            FA[i] = v2;
            FC[i] = FD;
            FB[i] = v3;
            h[i] = h[i - 1];
            TolArray[i] = TolArray[i - 1];
            S[i] = S1;
            L[i] = L[i - 1];
		}
	} while (i > 0 && i < 49);
	return Result2;
}

__declspec(dllexport) double WINAPI gt_DM(double MinAge, double MaxAge, double Tau, double DP, double UZtime)
{
    double Result=0.0, Lambda=0.0;
	if ((MaxAge == 0) | (DP == 0))
	{
		return Result;
	}
	if (MinAge >= UZtime)
	{
		//double a=MinAge-UZtime, b=MaxAge-UZtime;
		double k1 = (MaxAge - MinAge)/2, k2 = (MinAge + MaxAge)/2 - UZtime;
		Result+=k1*Coeff5[0]*DispInt(k1*Roots5[0]+k2,Tau,Lambda,DP);
		Result+=k1*Coeff5[1]*DispInt(k1*Roots5[1]+k2,Tau,Lambda,DP);
		Result+=k1*Coeff5[2]*DispInt(k1*Roots5[2]+k2,Tau,Lambda,DP);
		Result+=k1*Coeff5[3]*DispInt(k1*Roots5[3]+k2,Tau,Lambda,DP);
		Result+=k1*Coeff5[4]*DispInt(k1*Roots5[4]+k2,Tau,Lambda,DP);
	}
	return Result;
}
/*__declspec(dllexport) double WINAPI gt_DM(double MinAge, double MaxAge, double Tau, double DP, double UZtime)
{
    int i=1;
    double Aarray[50], TolArray[50], h[50], FA[50], FB[50], FC[50], S[50], L[50];
    double FD, FE, S1, S2, v1, v2, v3, v4, v5, v6, v7, v8, Result = 0.0, Tolerance = 1E-03,Lambda=0.0;
	
	if ((MaxAge == 0) | (DP == 0))
	{
		return Result;
	}
	if (MinAge > UZtime)
	{
		TolArray[i] = 10 * Tolerance;
		Aarray[i] = MinAge - UZtime;
		h[i] = (MaxAge - MinAge) / 2;
		FA[i] = DispInt(MinAge - UZtime, Tau, Lambda, DP);
		FC[i] = DispInt(MinAge - UZtime + h[i], Tau, Lambda, DP);
		FB[i] = DispInt(MaxAge - UZtime, Tau, Lambda, DP);
		S[i] = h[i] * (FA[i] + 4 * FC[i] + FB[i]) / 3;
		L[i] = 1;
		do
		{
			FD = DispInt(Aarray[i] + h[i] / 2, Tau, Lambda, DP);
			FE = DispInt(Aarray[i] + 3 * h[i] / 2, Tau, Lambda, DP);
			S1 = h[i] * (FA[i] + 4 * FD + FC[i]) / 6;
			S2 = h[i] * (FC[i] + 4 * FE + FB[i]) / 6;
			v1 = Aarray[i];
			v2 = FA[i];
			v3 = FC[i];
			v4 = FB[i];
			v5 = h[i];
			v6 = TolArray[i];
			v7 = S[i];
			v8 = L[i];
			i--;
			if (fabs(S1 + S2 - v7) < v6)
			{
				Result += (S1 + S2);
			}
			else if (i <= 47)
			{
				i++;
				Aarray[i] = v1 + v5;
				FA[i] = v3;
				FC[i] = FE;
				FB[i] = v4;
				h[i] = v5 / 2;
				TolArray[i] = v6 / 2;
				S[i] = S2;
				L[i] = v8 + 1;
				i++;
				Aarray[i] = v1;
				FA[i] = v2;
				FC[i] = FD;
				FB[i] = v3;
				h[i] = h[i - 1];
				TolArray[i] = TolArray[i - 1];
				S[i] = S1;
				L[i] = L[i - 1];
			}
		} while (i > 0 && i < 49); 
	}
	return Result;
}
*/
__declspec(dllexport) double WINAPI DM(
						FP DateRange[], FP TracerRange[], double Tau, double SampleDate, 
						double Lambda, double DP, double UZtime, LPXLOPER12 HeliumThree, LPXLOPER12 InitialTrit, LPXLOPER12 TritInitialTritRatio)
{
	double Result = 0.0;
	double DR, Cin; // pointers to Date Range and Tracer Input
	double EndDate, CinHe3, DMnoDecay, TimeIncrement, Lambda2, MaxAge=0.0, Integral, IntegralNoDecay, Change=1000, DMprev = 0.0, CummFrac=0.0;
	__int32 i,j, nIters, StepInc, StopCriteria;
	nIters = 1000000;
	if ((Tau == 0) | (DP == 0) | (SampleDate == 0) | (DateRange->rows != TracerRange->rows))
	{
		return Result;
	}
    if (DP < 1.0)
	{
        TimeIncrement = 1.0 / 12.0;
	}
    else
	{
        TimeIncrement = 1.0 / 12.0 / 2.0;
	}
    if (Tau >= 100.) // reset time increment and end date for large tau
	{
        TimeIncrement = fabs(DateRange->array[1]-DateRange->array[0]); //MinTimeInc(DateRange);
		if (fabs(DateRange->array[DateRange->rows - 1]-DateRange->array[DateRange->rows - 2]) < TimeIncrement)
		{
			TimeIncrement = fabs(DateRange->array[DateRange->rows - 1]-DateRange->array[DateRange->rows - 2]);
		}
	}
    EndDate = floor (SampleDate) + floor ((SampleDate-floor (SampleDate)) / TimeIncrement) * TimeIncrement - UZtime;
    if (EndDate == SampleDate - UZtime)
	{
        EndDate = EndDate - 1.0/12.0;
	}
    j = DateRange->rows-1;
	StepInc = 1;
	StopCriteria = 0;
	if (DateRange->array[j] < DateRange->array[0])
	{
		StepInc = -1;
		StopCriteria = j;
		j = 0;
	}
	DR = DateRange->array[j];
	//i=2;
	//MaxAge = i * TimeIncrement;
    if (HeliumThree->val.xbool == TRUE)
	{
        if (Lambda == 0.0)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		for (i = 2; i<= nIters; i++) //do
		{
            DMprev = Result;
			MaxAge = i * TimeIncrement;
            while ((EndDate - MaxAge) < DR && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					return Result;
				}
			}
            CinHe3 = Cin * exp(-Lambda2 * UZtime) * (1. - exp(-Lambda2 * (MaxAge - TimeIncrement / 2)));
            Integral = gt_DMint(MaxAge - TimeIncrement, MaxAge, Tau, 0.0, DP);
            Result += CinHe3 * Integral;
			//CummFrac+=Integral;
			if (MaxAge > Tau) 
			{
				Change = (Result - DMprev) / Result * 10;
				if (Change < Tol)
					return Result;
			}
			//i++;
		} //while (Change > Tol && i <= nIters);
	}
    else if (TritInitialTritRatio->val.xbool == TRUE)
	{
        if (Lambda == 0.0)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		DMnoDecay = 0.0;
		for (i = 2; i<= nIters; i++) //do
		{
			DMprev = DMnoDecay;
			MaxAge = i * TimeIncrement;
            while ((EndDate - MaxAge) < DR && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					break;
				}
			}
            IntegralNoDecay = gt_DMint(MaxAge - TimeIncrement, MaxAge, Tau, 0.0, DP);
            DMnoDecay += Cin * exp(-Lambda2 * UZtime) * IntegralNoDecay;
            Integral = gt_DMint(MaxAge - TimeIncrement, MaxAge, Tau, Lambda2, DP);
            Result += Cin * exp(-Lambda2 * UZtime) * Integral;
			//CummFrac+=IntegralNoDecay;
			if (MaxAge > Tau) 
			{
				Change = (DMnoDecay - DMprev) / DMnoDecay * 10;
				if (Change < Tol)
					break;
			}
		} //while (Change > Tol && i <= nIters);
        Result = Result / DMnoDecay;
	}
	else
	{
        if (InitialTrit->val.xbool == TRUE)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		for (i = 2; i<= nIters; i++) //do
		{
			DMprev = Result;
			MaxAge = i * TimeIncrement;
            while ((EndDate - MaxAge) < DR && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					return Result;
				}
			}
            Integral = gt_DMint(MaxAge - TimeIncrement, MaxAge, Tau, Lambda, DP);
            Result += Cin * exp(-Lambda2 * UZtime) * Integral;
			//CummFrac+=Integral/exp(-Lambda*(MaxAge-TimeIncrement/2));
			if (MaxAge > Tau) 
			{
				Change = (Result - DMprev) / Result * 10;
				if (Change < Tol)
					return Result;
			}

			//i++;
			
		} //while (Change > Tol && i <= nIters);
	}
	return Result;
}

__declspec(dllexport) double WINAPI DM_He4(
                        double Uppm, double THppm, double Porosity, double SedDensity, double Tau, double DP, double HeSolnRate)
{
	double Result=0.0;
	double Cin, DMold, TimeIncrement, Integral, MaxAge, Change = 1.0, CummFrac=0.0;
	__int32 i, SimYears, nIters;
    
	if ((Tau == 0) | (DP == 0))
	{
		return Result;
	}
    SimYears = 1000000;
	if (Tau < 100)
	{
		TimeIncrement = 1.0/12.0;
	}
	else
	{
		if (Tau < 1000)
		{
			TimeIncrement = 0.5;
		}
		else
		{
			TimeIncrement = Tau / 1000;
		}
	}
    nIters = (__int32) (SimYears / TimeIncrement);
    Result = 0.0;
	if (HeSolnRate != 0.0)
	{
		Cin = HeSolnRate;
	}
	else
	{
		Cin = SedDensity / Porosity * (Udecay * Uppm + THdecay * THppm);
	}
	i=2;
	do //for (i = 2; i<= nIters; i++)
	{
        DMold = Result;
        MaxAge = i * TimeIncrement;
		i++;
        Integral = gt_DMint(MaxAge - TimeIncrement, MaxAge, Tau, 0.0, DP);
        Result += Cin * MaxAge * Integral;
        //CummFrac+=Integral;
		if (MaxAge > Tau) 
				Change = (Result - DMold) / Result * 100;
	} while (Change > Tol && i <= nIters);
	return Result;
}

__declspec(dllexport) double WINAPI gt_PFM(double MinAge, double MaxAge, double Tau, double UZtime)
{
	double Result = 0.0, BeginAge;
	BeginAge = Tau + UZtime;
	if (MinAge < BeginAge && MaxAge >= BeginAge)
	{
		Result = 1.0;
	}
	else
	{
		Result = 0.0;
	}
	return Result;
}

__declspec(dllexport) double WINAPI gt_EMM(double MinAge, double MaxAge,
                double Tau, double UZtime)
{
	double Result=0.0;
	double n=1.0, TauRes, Tstar, BeginAge;

	TauRes = Tau / (1 + log (n));
	Tstar = log (n) * TauRes;
	BeginAge = Tstar + UZtime;
	if (MinAge <= BeginAge && MaxAge > BeginAge)
	{
		Result = n * (exp(-(BeginAge - UZtime) / TauRes) - exp(-(MaxAge - UZtime) / TauRes));
		return Result;
	}
	if (MinAge >= BeginAge)
	{
		Result = n * (exp(-(MinAge - UZtime) / TauRes) - exp(-(MaxAge - UZtime) / TauRes));
		return Result;
	}
	return Result;
}
/*
__declspec(dllexport) double WINAPI gt_PEM(double MinAge, double MaxAge,
                double Tau, double PEMratio, double UZtime)
{
	double Result=0.0;
	double n, TauRes, Tstar, BeginAge;
	n = PEMratio +1;
	TauRes = Tau / (1 + log (n));
	Tstar = log (n) * TauRes;
	BeginAge = Tstar + UZtime;
	if (MinAge <= BeginAge && MaxAge > BeginAge)
	{
		Result = n * (exp(-(BeginAge - UZtime) / TauRes) - exp(-(MaxAge - UZtime) / TauRes));
		return Result;
	}
	if (MinAge >= BeginAge)
	{
		Result = n * (exp(-(MinAge - UZtime) / TauRes) - exp(-(MaxAge - UZtime) / TauRes));
		return Result;
	}
	return Result;
}
*/
__declspec(dllexport) double WINAPI gt_PEM(double MinAge, double MaxAge, double Tau, double PEM_Uratio, double PEM_Lratio, double UZtime)
{
	double Result=0.0;
	double EndAge, BeginAge, nS, nU, nL, nStar;
	double TauRes, Tstar, pU, pS, pL, TauUpper, TauLower;
	nU = PEM_Uratio + 1;
	nL = PEM_Lratio + 1;
	if (PEM_Uratio == 0 && PEM_Lratio > 0)
	{
		nStar = 1 / PEM_Lratio + 1;
		nS = nStar;
		TauRes = Tau / (1 - (log (nL) + 1) / nL) / nStar;
		Tstar = log (nL) * TauRes;
		EndAge = Tstar + UZtime;
		BeginAge = UZtime;
	}
	else if (PEM_Lratio > PEM_Uratio && PEM_Lratio != 0)
	{
		nStar = 1 / PEM_Uratio + 1;
		pU = 1 / nStar;
		pL = 1 / nL;
		pS = 1 - pU - pL;
		nS = 1 / pS;
		TauUpper = (1 - log (nU) / nU - 1 / nU) * nStar * pU;
		TauLower = (log (nL) + 1) * pL;
		TauRes = pS * Tau / (1 - TauUpper - TauLower);
		Tstar = log (nU) * TauRes;
		BeginAge = Tstar + UZtime;
		Tstar = log (nL) * TauRes;
		EndAge = Tstar + UZtime;
	}
	else if (PEM_Uratio > 0 && PEM_Lratio == 0)
	{
		nS = nU;
		TauRes = Tau / (1 + log (nU));
		Tstar = log (nU) * TauRes;
		BeginAge = Tstar + UZtime;
		Tstar = log (nL) * TauRes;
		EndAge = 1e6;
	}
	else if (PEM_Uratio == 0 && PEM_Lratio == 0)
	{
		nS = 1;
		TauRes = Tau;
		BeginAge = UZtime;
		EndAge = 1e6;
	}
	else
	{
		return Result;
	}
	if (MinAge <= BeginAge && MaxAge > BeginAge)
	{
		Result = nS * (exp(-(BeginAge - UZtime) / TauRes) - exp(-(MaxAge - UZtime) / TauRes));
		return Result;
	}
	else if (MaxAge > BeginAge && MaxAge <= EndAge)
	{
		Result = nS * (exp(-(MinAge - UZtime) / TauRes) - exp(-(MaxAge - UZtime) / TauRes));
		return Result;
	}
	else if (MinAge <= EndAge && MaxAge > EndAge)
	{
		Result = nS * (exp(-(MinAge - UZtime) / TauRes) - exp(-(EndAge - UZtime) / TauRes));
		return Result;
	}
	return Result;
}
__declspec(dllexport) double WINAPI gt_EPM(double MinAge, double MaxAge,
                double Tau, double EPMratio, double UZtime)
{
	double Result = 0.0;
	double n, BeginAge;
	n = EPMratio +1;
	BeginAge = Tau * (1 - 1 / n) + UZtime;
	if (MinAge <= BeginAge && MaxAge > BeginAge)
	{
		Result = exp(-(BeginAge - UZtime) * n / Tau + n - 1) - exp(-(MaxAge - UZtime) * n / Tau + n - 1);
		return Result;
	}
	if (MinAge >= BeginAge)
	{
		Result = exp(-(MinAge - UZtime) * n / Tau + n - 1) - exp(-(MaxAge - UZtime) * n / Tau + n - 1);
		return Result;
	}
	return Result;
}
double gt_PFMint(double MinAge, double MaxAge, double Tau, double UZtime)
{
	double Result = 0.0, BeginAge;
	BeginAge = Tau + UZtime;
	if (MinAge < BeginAge && MaxAge >= BeginAge)
	{
		Result = 1.0;
	}
	else
	{
		Result = 0.0;
	}
	return Result;
}

double gt_EMMint(double MinAge, double MaxAge, double Tau, double UZtime)
{
	double Result=0.0;
	double n=1.0, TauRes, Tstar, BeginAge;

	TauRes = Tau / (1 + log (n));
	Tstar = log (n) * TauRes;
	BeginAge = Tstar + UZtime;
	if (MinAge <= BeginAge && MaxAge > BeginAge)
	{
		Result = n * (exp(-(BeginAge - UZtime) / TauRes) - exp(-(MaxAge - UZtime) / TauRes));
		return Result;
	}
	if (MinAge >= BeginAge)
	{
		Result = n * (exp(-(MinAge - UZtime) / TauRes) - exp(-(MaxAge - UZtime) / TauRes));
		return Result;
	}
	return Result;
}
__declspec(dllexport) double WINAPI gt_GAM(double MinAge, double MaxAge, double Tau, double Alpha, double UZtime)
{
	double BeginAge, Beta, Result1=0.0, Result2=0.0;

	BeginAge = UZtime;
	Beta = Tau/Alpha;
	
	if (MinAge <= BeginAge && MaxAge > BeginAge)
	{
		Result1 = ReturnExcelGamma(BeginAge-UZtime, Alpha, Beta);
		Result2 = ReturnExcelGamma(MaxAge-UZtime, Alpha, Beta);
	}
	else if (MinAge >= BeginAge)
	{	
		Result1 = ReturnExcelGamma(MinAge-UZtime, Alpha, Beta);
		Result2 = ReturnExcelGamma(MaxAge-UZtime, Alpha, Beta);
	}

	return Result2-Result1;
}
double WINAPI gt_GAMint(double MinAge, double MaxAge, double Tau, double Alpha, double UZtime)
{
	double BeginAge, Beta, Result1=0.0, Result2=0.0;

	BeginAge = UZtime;
	Beta = Tau/Alpha;
	
	if (MinAge <= BeginAge && MaxAge > BeginAge)
	{
		Result1 = ReturnExcelGamma(BeginAge-UZtime, Alpha, Beta);
		Result2 = ReturnExcelGamma(MaxAge-UZtime, Alpha, Beta);
	}
	else if (MinAge >= BeginAge)
	{	
		Result1 = ReturnExcelGamma(MinAge-UZtime, Alpha, Beta);
		Result2 = ReturnExcelGamma(MaxAge-UZtime, Alpha, Beta);
	}

	return Result2-Result1;
}

/*double gt_PEMint(double MinAge, double MaxAge, double Tau, double PEMratio, double UZtime)
{
	double Result=0.0;
	double n, TauRes, Tstar, BeginAge;
	n = PEMratio +1;
	TauRes = Tau / (1 + log (n));
	Tstar = log (n) * TauRes;
	BeginAge = Tstar + UZtime;
	if (MinAge <= BeginAge && MaxAge > BeginAge)
	{
		Result = n * (exp(-(BeginAge - UZtime) / TauRes) - exp(-(MaxAge - UZtime) / TauRes));
		return Result;
	}
	if (MinAge >= BeginAge)
	{
		Result = n * (exp(-(MinAge - UZtime) / TauRes) - exp(-(MaxAge - UZtime) / TauRes));
		return Result;
	}
	return Result;
}
*/
double gt_PEMint(double MinAge, double MaxAge, double Tau, double PEM_Uratio, double PEM_Lratio, double UZtime)
{
	double Result=0.0;
	double EndAge, BeginAge, nS, nU, nL, nStar;
	double TauRes, Tstar, pU, pS, pL, TauUpper, TauLower;
	nU = PEM_Uratio + 1;
	nL = PEM_Lratio + 1;
	if (PEM_Uratio == 0 && PEM_Lratio > 0)
	{
		nStar = 1 / PEM_Lratio + 1;
		nS = nStar;
		TauRes = Tau / (1 - (log (nL) + 1)/nL)*nStar; 
		Tstar = log (nL) * TauRes;
		EndAge = Tstar + UZtime;
		BeginAge = UZtime;
	}
	else if (PEM_Lratio > PEM_Uratio && PEM_Lratio != 0)
	{
		nStar = 1 / PEM_Lratio + 1;
		pU = 1 / nStar;
		pL = 1 / nL;
		pS = 1 - pU - pL;
		nS = 1 / pS;
		TauUpper = (1 - log (nU) / nU - 1 / nU) * nStar * pU;
		TauLower = (log (nL) + 1) * pL;
		TauRes = pS * Tau / (1 - TauUpper - TauLower);
		Tstar = log (nU) * TauRes;
		BeginAge = Tstar + UZtime;
		Tstar = log (nL) * TauRes;
		EndAge = Tstar + UZtime;
	}
	else if (PEM_Uratio > 0 && PEM_Lratio == 0)
	{
		nS = nU;
		TauRes = Tau / (1 + log (nU));
		Tstar = log (nU) * TauRes;
		BeginAge = Tstar + UZtime;
		Tstar = log (nL) * TauRes;
		EndAge = 1e6;
	}
	else if (PEM_Uratio == 0 && PEM_Lratio == 0)
	{
		nS = 1;
		TauRes = Tau;
		BeginAge = UZtime;
		EndAge = 1e6;
	}
	else
	{
		return Result;
	}
	if (MinAge <= BeginAge && MaxAge > BeginAge)
	{
		Result = nS * (exp(-(BeginAge - UZtime) / TauRes) - exp(-(MaxAge - UZtime) / TauRes));
		return Result;
	}
	else if (MaxAge > BeginAge && MaxAge <= EndAge)
	{
		Result = nS * (exp(-(MinAge - UZtime) / TauRes) - exp(-(MaxAge - UZtime) / TauRes));
		return Result;
	}
	else if (MinAge <= EndAge && MaxAge > EndAge)
	{
		Result = nS * (exp(-(MinAge - UZtime) / TauRes) - exp(-(EndAge - UZtime) / TauRes));
		return Result;
	}
	return Result;
}
double gt_EPMint(double MinAge, double MaxAge, double Tau, double EPMratio, double UZtime)
{
	double Result = 0.0;
	double n, BeginAge;
	n = EPMratio +1;
	BeginAge = Tau * (1 - 1 / n) + UZtime;
	if (MinAge <= BeginAge && MaxAge > BeginAge)
	{
		Result = exp(-(BeginAge - UZtime) * n / Tau + n - 1) - exp(-(MaxAge - UZtime) * n / Tau + n - 1);
		return Result;
	}
	if (MinAge >= BeginAge)
	{
		Result = exp(-(MinAge - UZtime) * n / Tau + n - 1) - exp(-(MaxAge - UZtime) * n / Tau + n - 1);
		return Result;
	}
	return Result;
}
__declspec(dllexport) double WINAPI gt_BMM_PFM(double MinAge, double MaxAge,
                double Tau, double UZtime, double MixingFraction, double SecondFraction)
{
	double Result = 0.0, FirstFraction;
	FirstFraction = gt_PFMint(MinAge, MaxAge, Tau, UZtime);
	Result = (1 - MixingFraction) * SecondFraction + (MixingFraction * FirstFraction);
	return Result;
}
__declspec(dllexport) double WINAPI gt_BMM_EMM(double MinAge, double MaxAge,
                double Tau, double UZtime, double MixingFraction, double SecondFraction)
{
	double Result = 0.0, FirstFraction;
	FirstFraction = gt_EMMint(MinAge, MaxAge, Tau, UZtime);
	Result = (1 - MixingFraction) * SecondFraction + (MixingFraction * FirstFraction);
	return Result;
}
__declspec(dllexport) double WINAPI gt_BMM_PEM(double MinAge, double MaxAge,
                double Tau, double PEM_Uratio, double PEM_Lratio, double UZtime, double MixingFraction, double SecondFraction)
{
	double Result = 0.0, FirstFraction;
	FirstFraction = gt_PEMint(MinAge, MaxAge, Tau, PEM_Uratio, PEM_Lratio, UZtime);
	Result = (1 - MixingFraction) * SecondFraction + (MixingFraction * FirstFraction);
	return Result;
}
__declspec(dllexport) double WINAPI gt_BMM_GAM(double MinAge, double MaxAge,
                double Tau, double Alpha, double UZtime, double MixingFraction, double SecondFraction)
{
	double Result = 0.0, FirstFraction;
	FirstFraction = gt_GAMint(MinAge, MaxAge, Tau, Alpha, UZtime);
	Result = (1 - MixingFraction) * SecondFraction + (MixingFraction * FirstFraction);
	return Result;
}
__declspec(dllexport) double WINAPI gt_BMM_EPM(double MinAge, double MaxAge,
                double Tau, double EPMratio, double UZtime, double MixingFraction, double SecondFraction)
{
	double Result = 0.0, FirstFraction;
	FirstFraction = gt_EPMint(MinAge, MaxAge, Tau, EPMratio, UZtime);
	Result = (1 - MixingFraction) * SecondFraction + (MixingFraction * FirstFraction);
	return Result;
}
__declspec(dllexport) double WINAPI gt_BMM_DM(double MinAge, double MaxAge,
                double Tau, double DP, double UZtime, double MixingFraction, double SecondFraction)
{
	double Result = 0.0, FirstFraction;
	FirstFraction = gt_DM(MinAge, MaxAge, Tau, DP, UZtime);
	Result = (1 - MixingFraction) * SecondFraction + (MixingFraction * FirstFraction);
	return Result;
}
__declspec(dllexport) double WINAPI BMM_DM(
						FP DateRange[], FP TracerRange[], double Tau, double SampleDate, 
						double Lambda, double DP, double UZtime, double MixingFraction, double Cold, LPXLOPER12 HeliumThree, 
						LPXLOPER12 InitialTrit, LPXLOPER12 TritInitialTritRatio, double DICyoung, double DICold)
{
	double Result, Cyoung, C14mix;

	Result = 0.0;
	Cyoung = DM(DateRange, TracerRange, Tau, SampleDate, Lambda,DP, UZtime,HeliumThree,InitialTrit,TritInitialTritRatio);
	if (DICyoung == 0 || DICold == 0)
	{
		Result = (1 - MixingFraction) * Cold + (MixingFraction * Cyoung);
	}
	else
	{
		C14mix = (1 - MixingFraction) * DICold + MixingFraction * DICyoung;
		Result = ((1 - MixingFraction) * Cold * DICold + (MixingFraction * Cyoung * DICyoung)) / C14mix;
	}
	return Result;
}
__declspec(dllexport) double WINAPI BMM_EMM(
						FP DateRange[], FP TracerRange[], double Tau, double SampleDate, 
						double Lambda, double UZtime, double MixingFraction, double Cold, LPXLOPER12 HeliumThree, 
						LPXLOPER12 InitialTrit, LPXLOPER12 TritInitialTritRatio, double DICyoung, double DICold)
{
	double Result, Cyoung, C14mix;
	Result = 0.0;
	Cyoung = EMM(DateRange, TracerRange,Tau, SampleDate, Lambda, UZtime,HeliumThree,InitialTrit,TritInitialTritRatio);
	if (DICyoung == 0 || DICold == 0)
	{
		Result = (1 - MixingFraction) * Cold + (MixingFraction * Cyoung);
	}
	else
	{
		C14mix = (1 - MixingFraction) * DICold + MixingFraction * DICyoung;
		Result = ((1 - MixingFraction) * Cold * DICold + (MixingFraction * Cyoung * DICyoung)) / C14mix;
	}
	return Result;
}
__declspec(dllexport) double WINAPI BMM_PEM(
						FP DateRange[], FP TracerRange[], double Tau, double SampleDate, 
						double Lambda, double PEM_Uratio, double PEM_Lratio, double UZtime, double MixingFraction, double Cold, LPXLOPER12 HeliumThree, 
						LPXLOPER12 InitialTrit, LPXLOPER12 TritInitialTritRatio, double DICyoung, double DICold)
{
	double Result, Cyoung, C14mix;
	Result = 0.0;
	Cyoung = PEM(DateRange, TracerRange,Tau, SampleDate, Lambda,PEM_Uratio, PEM_Lratio, UZtime,HeliumThree,InitialTrit,TritInitialTritRatio);
	if (DICyoung == 0 || DICold == 0)
	{
		Result = (1 - MixingFraction) * Cold + (MixingFraction * Cyoung);
	}
	else
	{
		C14mix = (1 - MixingFraction) * DICold + MixingFraction * DICyoung;
		Result = ((1 - MixingFraction) * Cold * DICold + (MixingFraction * Cyoung * DICyoung)) / C14mix;
	}
	return Result;
}
__declspec(dllexport) double WINAPI BMM_EPM(
						FP DateRange[], FP TracerRange[], double Tau, double SampleDate, 
						double Lambda, double EPMratio, double UZtime, double MixingFraction, double Cold, LPXLOPER12 HeliumThree, 
						LPXLOPER12 InitialTrit, LPXLOPER12 TritInitialTritRatio, double DICyoung, double DICold)
{
	double Result, Cyoung, C14mix;
	Result = 0.0;
	Cyoung = EPM(DateRange, TracerRange,Tau, SampleDate, Lambda,EPMratio, UZtime,HeliumThree,InitialTrit,TritInitialTritRatio);
	if (DICyoung == 0 || DICold == 0)
	{
		Result = (1 - MixingFraction) * Cold + (MixingFraction * Cyoung);
	}
	else
	{
		C14mix = (1 - MixingFraction) * DICold + MixingFraction * DICyoung;
		Result = ((1 - MixingFraction) * Cold * DICold + (MixingFraction * Cyoung * DICyoung)) / C14mix;
	}
	return Result;
}
__declspec(dllexport) double WINAPI BMM_GAM(
						FP DateRange[], FP TracerRange[], double Tau, double SampleDate, 
						double Lambda, double Alpha, double UZtime, double MixingFraction, double Cold, bool HeliumThree, 
						bool InitialTrit, bool TritInitialTritRatio, double DICyoung, double DICold)
{
	double Result, Cyoung, C14mix;
	Result = 0.0;
	Cyoung = GAM(DateRange, TracerRange,Tau, SampleDate, Lambda,Alpha, UZtime,HeliumThree,InitialTrit,TritInitialTritRatio);
	if (DICyoung == 0 || DICold == 0)
	{
		Result = (1 - MixingFraction) * Cold + (MixingFraction * Cyoung);
	}
	else
	{
		C14mix = (1 - MixingFraction) * DICold + MixingFraction * DICyoung;
		Result = ((1 - MixingFraction) * Cold * DICold + (MixingFraction * Cyoung * DICyoung)) / C14mix;
	}
	return Result;
}
__declspec(dllexport) double WINAPI BMM_PFM(
						FP DateRange[], FP TracerRange[], double Tau, double SampleDate, 
						double Lambda, double UZtime, double MixingFraction, double Cold, LPXLOPER12 HeliumThree, 
						LPXLOPER12 InitialTrit, LPXLOPER12 TritInitialTritRatio, double DICyoung, double DICold)
{
	double Result, Cyoung, C14mix;
	Result = 0.0;
	Cyoung = PFM(DateRange, TracerRange,Tau, SampleDate, Lambda, UZtime,HeliumThree,InitialTrit,TritInitialTritRatio);
	if (DICyoung == 0 || DICold == 0)
	{
		Result = (1 - MixingFraction) * Cold + (MixingFraction * Cyoung);
	}
	else
	{
		C14mix = (1 - MixingFraction) * DICold + MixingFraction * DICyoung;
		Result = ((1 - MixingFraction) * Cold * DICold + (MixingFraction * Cyoung * DICyoung)) / C14mix;
	}
	return Result;
}

__declspec(dllexport) double WINAPI BMM_DM_He4(
						double Uppm, double THppm, double Porosity, double SedDensity, double Tau,
						double DP, double HeSolnRate, double MixingFraction, double Cold)
{
	double Result, Cyoung;
	Result = 0.0;
	Cyoung = DM_He4(Uppm, THppm, Porosity, SedDensity, Tau, DP, HeSolnRate);
	Result = (1 - MixingFraction) * Cold + (MixingFraction * Cyoung);
	return Result;
}
__declspec(dllexport) double WINAPI BMM_EMM_He4(
						double Uppm, double THppm, double Porosity, double SedDensity, double Tau,
						double HeSolnRate, double MixingFraction, double Cold)
{
	double Result, Cyoung;
	Result = 0.0;
	Cyoung = EMM_He4(Uppm, THppm, Porosity, SedDensity, Tau, HeSolnRate);
	Result = (1 - MixingFraction) * Cold + (MixingFraction * Cyoung);
	return Result;
}
__declspec(dllexport) double WINAPI BMM_PEM_He4(
						double Uppm, double THppm, double Porosity, double SedDensity, double Tau,
						double PEM_Uratio, double PEM_Lratio, double HeSolnRate, double MixingFraction, double Cold)
{
	double Result, Cyoung;
	Result = 0.0;
	Cyoung = PEM_He4(Uppm, THppm, Porosity, SedDensity, Tau, PEM_Uratio, PEM_Lratio, HeSolnRate);
	Result = (1 - MixingFraction) * Cold + (MixingFraction * Cyoung);
	return Result;
}
__declspec(dllexport) double WINAPI BMM_EPM_He4(
						double Uppm, double THppm, double Porosity, double SedDensity, double Tau,
						double EPMratio, double HeSolnRate, double MixingFraction, double Cold)
{
	double Result, Cyoung;
	Result = 0.0;
	Cyoung = EPM_He4(Uppm, THppm, Porosity, SedDensity, Tau, EPMratio, HeSolnRate);
	Result = (1 - MixingFraction) * Cold + (MixingFraction * Cyoung);
	return Result;
}
__declspec(dllexport) double WINAPI BMM_GAM_He4(
						double Uppm, double THppm, double Porosity, double SedDensity, double Tau,
						double Alpha, double HeSolnRate, double MixingFraction, double Cold)
{
	double Result, Cyoung;
	Result = 0.0;
	Cyoung = GAM_He4(Uppm, THppm, Porosity, SedDensity, Tau, HeSolnRate,Alpha);
	Result = (1 - MixingFraction) * Cold + (MixingFraction * Cyoung);
	return Result;
}
__declspec(dllexport) double WINAPI BMM_PFM_He4(
						double Uppm, double THppm, double Porosity, double SedDensity, double Tau,
						double HeSolnRate, double MixingFraction, double Cold)
{
	double Result, Cyoung;
	Result = 0.0;
	Cyoung = PFM_He4(Uppm, THppm, Porosity, SedDensity, Tau, HeSolnRate);
	Result = (1 - MixingFraction) * Cold + (MixingFraction * Cyoung);
	return Result;
}


__declspec(dllexport) double WINAPI AlphaDensity(double T, double X, double Alpha, double Beta, double C, double Mu)
{
	double Phi,Result = 0.0;
	complex <double> c3, c5,c7;

	//
	if (Alpha==1)
	{
		double Tt = fabs(T);
		if (Tt==0.0)
			Tt=1e-08;
		Phi = -(2 / PI) * log (fabs(Tt));
	}
	else
	{
		Phi = tan (PI * Alpha / 2);
	}
	complex <double> c1(1,-Beta*sign(T)*Phi);
	c3 = pow (C*fabs(T),Alpha)*c1;
	complex <double> c4(0,T*Mu);
	c5 = exp(c4-c3);
	complex <double> c6(0,-X*T);
	c7=c5*exp(c6);
	return c7.real();
	/*c1.imag(-1);
	c1.real(0);
	
	
	c2 = imag ( c1 * (sign (T) * Phi * Beta));
	c3.imag(1);
	c3.real(0);
	if (c2 == 0)
	{
		c7=pow(C*fabs(T),Alpha);
		c8=c3*(T*Mu);
		c4 = exp (c8 - c7);
	}
	else
	{
		c5 = (1,c2);
		c4 = exp ((c3 * (T *Mu)) - pow(fabs(C * T), Alpha) * c5);
	}
	c6 = exp (c1 * (X * T)) * c4;
	Result = c6.real();
    return Result;*/
}

static int TerminationCode, NumOfLevels;

double AlphaStablePDFaux2(double a, double b, double fa, double fb, double IS, double X, double Alpha, double Beta, double C, double Mu) 
{                 
  double m = (a + b)/2, h = (b - a)/2, z = 2.0/3.0;
  double al = sqrt(z), beta=1/sqrt(double (5));
  double mll = m - al*h, ml = m - beta*h, mr = m + beta*h, mrr = m + al*h;
  double fmll = AlphaDensity(mll, X, Alpha,Beta,C,Mu), fml = AlphaDensity(ml, X, Alpha,Beta,C,Mu), fm = AlphaDensity(m, X, Alpha,Beta,C,Mu);
  double fmr = AlphaDensity(mr, X, Alpha,Beta,C,Mu), fmrr = AlphaDensity(mrr, X, Alpha,Beta,C,Mu);
  double i2 = (h/6)*(fa + fb + 5*(fml + fmr));
  double i1 = h/1470*(77*(fa + fb)+432*(fmll + fmrr)+625*(fml + fmr)+672*fm);
  double Eval = IS + (i1 - i2);
  if ((Eval == IS) || (mll <= a) || (b < mrr))
  {
	  //if ((m <= a || b < m) && TerminationCode==0)
		 // TerminationCode=1;
	  return i1;
  }
  else
  {
	  return AlphaStablePDFaux2(a, mll, fa, fmll, IS, X, Alpha, Beta, C, Mu) +                    
		  AlphaStablePDFaux2(mll, ml, fmll, fml, IS, X, Alpha, Beta, C, Mu) +
		  AlphaStablePDFaux2(ml, m, fml, fm, IS, X, Alpha, Beta, C, Mu) +
		  AlphaStablePDFaux2(m, mr, fm, fmr, IS, X, Alpha, Beta, C, Mu) +
		  AlphaStablePDFaux2(mr, mrr, fmr, fmrr, IS, X, Alpha, Beta, C, Mu) +
		  AlphaStablePDFaux2(mrr, b, fmrr, fb, IS, X, Alpha, Beta, C, Mu);
  }                    
}         
 
__declspec(dllexport) double WINAPI AlphaStablePDF(double X, double Alpha, double Beta, double C, double Mu)
{
	double Result=0.0,a=-500, b=500, IS=0.0, epsilon=2.22e-016,Tol = 1e-07,Ys[13];
	int k;
	double m = (a + b)/2, h = (b - a)/2, z = 2.0 / 3.0;
	double al = sqrt(z), beta=1/sqrt(double (5)); 
	double Xs[13] = {a, m-glX[0]*h, m-al*h, m-glX[1]*h, m-beta*h, m-glX[2]*h, m, m+glX[2]*h,
		m+beta*h, m+glX[1]*h, m+al*h, m+glX[0]*h, b};
	//double fa = AlphaDensity(a,X,Alpha, Beta, C, Mu), fb = AlphaDensity(b, X, Alpha, Beta, C, Mu), fm = AlphaDensity(m, X, Alpha, Beta, C, Mu);                                                           
	for (k = 0; k< 13; k++)
	{
		Ys[k] = AlphaDensity(Xs[k],X,Alpha, Beta, C, Mu);
	}
	double fa = Ys[0], fb = Ys[12];
	double i2 = h/6*(Ys[0]+Ys[12]+5*(Ys[4]+Ys[8]));
	double i1 = h/1470*(77*(Ys[0]+Ys[12])+432*(Ys[2]+Ys[10])+625*(Ys[4]+Ys[8])+672*Ys[6]);
	IS = h*glA[0]*(Ys[0]+Ys[12])+glA[1]*(Ys[1]+Ys[11])+glA[2]*(Ys[2]+Ys[10])+
		glA[3]*(Ys[3]+Ys[9])+glA[4]*(Ys[4]+Ys[8])+glA[5]*(Ys[5]+Ys[7])+glA[6]*Ys[6];
	int S = sign(IS);
	if (S==0)
		S = 1;
	double erri1 = fabs(i1-IS);
	double erri2 = fabs(i2-IS);
	int R=1;
	if (erri2 !=0)
		R = int(erri1 / erri2);
	if (R > 0 && R < 1)
		Tol = Tol/R;
	IS = S*fabs(IS)*Tol/epsilon;
	if (IS == 0)
		IS = b - a;
	Result = AlphaStablePDFaux2(a, b, fa, fb, IS, X, Alpha,Beta,C,Mu);
	return Result/(2*PI);
}

inline double AlphaDensityInt(double T, double& X, double& Alpha)
{

	//double Test = cos(X*T) * exp(-pow(fabs(T), Alpha));
	//complex <double> c1(0,-X * T);
	//complex <double> c4(exp(-pow(fabs(T), Alpha)),0);
	//complex <double> c6 = exp (c1) * c4;
    //return c6.real();
	return cos(X*T) * exp(-pow(fabs(T), Alpha));
}
 double adaptiveSimpsonsAux(double a, double b, double epsilon, double S, double fa, double fb, double fc, int bottom, double& X, double& Alpha) 
{                 
  double c = (a + b)/2, h = b - a;                                                                  
  double d = (a + c)/2, e = (c + b)/2;                                                              
  double fd = AlphaDensityInt(d, X, Alpha), fe = AlphaDensityInt(e, X, Alpha);                                                                      
  double Sleft = (h/12)*(fa + 4*fd + fc);                                                           
  double Sright = (h/12)*(fc + 4*fe + fb);                                                          
  double S2 = Sleft + Sright;                                                                       
  if (bottom <= 0 || fabs(S2 - S) <= 15*epsilon)                                                    
    return S2 + (S2 - S)/15;       
  NumOfLevels++;
  return adaptiveSimpsonsAux(a, c, epsilon/2, Sleft,  fa, fc, fd, bottom-1, X, Alpha) +                    
         adaptiveSimpsonsAux(c, b, epsilon/2, Sright, fc, fb, fe, bottom-1, X, Alpha);                     
}         

double adaptiveSimpsons(double& T, double& Alpha, double& Tau, double& DP, double& Lambda)
{   // recursion cap        
	double a=0.0, b=500, epsilon=1e-08,IntCriteria,Result=0.0;
	double Exponent, Dispersion, Velocity, X, Var, F1, F2, F3;
	int maxRecursionDepth=100;
	Exponent=1/Alpha;
	Velocity=1/Tau;
	Dispersion=DP*Velocity;
	Var=abs(cosf(PI*.5*Alpha));
	F1=pow(T,Exponent);
	F2=fabs(1-T/Tau);
	F3=pow((DP*Var/Tau),Exponent);
	X=F2/(F1*F3);
	double c = (a + b)/2, h = b - a;                        
	if (Alpha == 2.0)
		IntCriteria = 14;
	else
		//IntCriteria = 10500*pow(Alpha,-4.626);
		IntCriteria = 10000*pow(Alpha,-4.534);
	if (X < IntCriteria)
	{
		double fa = AlphaDensityInt(a,X,Alpha), fb = AlphaDensityInt(b, X, Alpha), fc = AlphaDensityInt(c, X, Alpha);                                                           
		double S = (h/6)*(fa + 4*fc + fb);                                                                
		Result = adaptiveSimpsonsAux(a, b, epsilon, S, fa, fb, fc, maxRecursionDepth, X, Alpha);
	}
	double Sigma=PI*(pow((Var*Dispersion*T),Exponent)*T);
	return Result/Sigma*exp (-Lambda*T);
}
static const double Sqrt1=0.81649658092772603, Sqrt2=0.44721359549995793;
 double AlphaStablePDFauxGL(double a, double b, double fa, double fb, double IS, double& X, double& Alpha) 
{                 
  double m = (a + b)/2, h = (b - a)/2;
  double mll = m - Sqrt1*h, ml = m - Sqrt2*h, mr = m + Sqrt2*h, mrr = m + Sqrt1*h;
  double fmll = AlphaDensityInt(mll, X, Alpha), fml = AlphaDensityInt(ml, X, Alpha), fm = AlphaDensityInt(m, X, Alpha);
  double fmr = AlphaDensityInt(mr, X, Alpha), fmrr = AlphaDensityInt(mrr, X, Alpha);
  double i2 = (h/6)*(fa + fb + 5*(fml + fmr));
  double i1 = h/1470*(77*(fa + fb)+432*(fmll + fmrr)+625*(fml + fmr)+672*fm);
  double Eval = IS + (i1 - i2);
  if ((Eval == IS) || (mll <= a) || (b < mrr))
  {
	  //if ((m <= a || b < m) && TerminationCode==0)
		//  TerminationCode=1;
	  return i1;
  }
  else
  {
	  return AlphaStablePDFauxGL(a, mll, fa, fmll, IS, X, Alpha) +                    
		  AlphaStablePDFauxGL(mll, ml, fmll, fml, IS, X, Alpha) +
		  AlphaStablePDFauxGL(ml, m, fml, fm, IS, X, Alpha) +
		  AlphaStablePDFauxGL(m, mr, fm, fmr, IS, X, Alpha) +
		  AlphaStablePDFauxGL(mr, mrr, fmr, fmrr, IS, X, Alpha) +
		  AlphaStablePDFauxGL(mrr, b, fmrr, fb, IS, X, Alpha);
  }
} 

 double AlphaStablePDFInt(double T, double& Alpha, double& Tau, double& DP, double& Lambda)
 {
	 //Computes the integral of an aplha stable density (with Beta = 0, Mu=0) using
	 //adaptive gauss-lobatto numerical integration (Gander and Gautschi, 2000; Adaptive Quadrature -- Revisited)
	 //with modification

	 //Modified so that the calculation is one sided (positive X) with a = 0.0 and b is dependent on alpha

	double Result=0.0,a=0.0, b=561*pow(Alpha, -4.583), IS=0.0, epsilon=2.22e-016,Tol = 1e-08;
	double Exponent=1/Alpha, Velocity=1/Tau, Dispersion=DP*Velocity, Var=abs(cosf(PI*0.5*Alpha));
	double F1=pow(T,Exponent), F2=abs(1-T/Tau), F3=pow((DP*Var/Tau),Exponent), X=F2/(F1*F3);
	double Sigma=PI*(pow((Var*Dispersion*T),Exponent)*T), IntCriteria = 10000*pow(Alpha,-4.534);
	if (Alpha == 2.0)
		IntCriteria = 14;
	if (X < IntCriteria)
	{
		//double m = (a + b)/2, h = (b - a)/2;
		double m = b/2, h = m;
		//double Xs[13] = {a, m-glX[0]*h, m-al*h, m-glX[1]*h, m-Sqrt2*h, m-glX[2]*h, m, m+glX[2]*h,
		//	m+Sqrt2*h, m+glX[1]*h, m+al*h, m+glX[0]*h, b};
		double Ys[13] = {AlphaDensityInt(a,X,Alpha), 
			AlphaDensityInt(m-glX[0]*h,X,Alpha), 
			AlphaDensityInt(m-Sqrt1*h,X,Alpha),
			AlphaDensityInt(m-glX[1]*h,X,Alpha),
			AlphaDensityInt(m-Sqrt2*h,X,Alpha),
			AlphaDensityInt(m-glX[2]*h,X,Alpha),
			AlphaDensityInt(m,X,Alpha),
			AlphaDensityInt(m+glX[2]*h,X,Alpha),
			AlphaDensityInt(m+Sqrt2*h,X,Alpha),
			AlphaDensityInt(m+glX[1]*h,X,Alpha),
			AlphaDensityInt(m+Sqrt1*h,X,Alpha),
			AlphaDensityInt(m+glX[0]*h,X,Alpha),
			AlphaDensityInt(b,X,Alpha)};
		double fa = Ys[0], fb = Ys[12];
		double i2 = h/6*(Ys[0]+Ys[12]+5*(Ys[4]+Ys[8]));
		double i1 = h/1470*(77*(Ys[0]+Ys[12])+432*(Ys[2]+Ys[10])+625*(Ys[4]+Ys[8])+672*Ys[6]);
		IS = h*glA[0]*(Ys[0]+Ys[12])+glA[1]*(Ys[1]+Ys[11])+glA[2]*(Ys[2]+Ys[10])+
			glA[3]*(Ys[3]+Ys[9])+glA[4]*(Ys[4]+Ys[8])+glA[5]*(Ys[5]+Ys[7])+glA[6]*Ys[6];
		int S = sign(IS);
		if (S==0)
			S = 1;
		double erri1 = fabs(i1-IS);
		double erri2 = fabs(i2-IS);
		double R=1.0;
		if (erri2 !=0)
			R = erri1 / erri2;
		if (R > 0.0 && R < 1.0)
			Tol = Tol/R;
		IS = S*fabs(IS)*Tol/epsilon;
		if (IS == 0)
			//IS = b - a;
			IS = b;
		Result = AlphaStablePDFauxGL(a, b, fa, fb, IS, X, Alpha);
	}
	return Result/Sigma*exp (-Lambda*T);
 }

  double gtFDMaux(double a, double b, double epsilon, double S, double fa, double fb, double fc, int bottom, double Alpha, double Tau, double DP) 
{                 
  double c = (a + b)/2, h = b - a;                                                                  
  double d = (a + c)/2, e = (c + b)/2;
  double Lambda = 0.0;
  double fd = AlphaStablePDFInt(d, Alpha,Tau,DP,Lambda), fe = AlphaStablePDFInt(e, Alpha,Tau,DP,Lambda);
  double Sleft = (h/12)*(fa + 4*fd + fc);                                                           
  double Sright = (h/12)*(fc + 4*fe + fb);                                                          
  double S2 = Sleft + Sright;                                                                       
  if (bottom <= 0 || fabs(S2 - S) <= 15*epsilon)                                                    
    return S2 + (S2 - S)/15;       
  NumOfLevels++;
  return gtFDMaux(a, c, epsilon/2, Sleft,  fa, fc, fd, bottom-1, Alpha,Tau,DP) +                    
         gtFDMaux(c, b, epsilon/2, Sright, fc, fb, fe, bottom-1, Alpha,Tau,DP); 
} 

__declspec(dllexport) double WINAPI gt_FDM(double MinAge, double MaxAge, double Alpha, double Tau, double DP, double UZtime)
{	
	double Result=0.0;
	if ((MaxAge <= 0) | (DP <= 0) | (Alpha <= 1) | (Alpha > 2) | (Tau <= 0))
	{
		return Result;
	}
	if (MinAge > UZtime)
	{
		int maxRecursionDepth=10;
		double a=MinAge-UZtime, b=MaxAge-UZtime, Tol = 1e-04;                                                                                                                          
		double c = (a + b)/2, h = b - a;                        
		TerminationCode=0, NumOfLevels=0;
		double Lambda=0.0;
		double fa = AlphaStablePDFInt(a,Alpha,Tau,DP,Lambda), fb = AlphaStablePDFInt(b, Alpha, Tau,DP, Lambda), fc = AlphaStablePDFInt(c, Alpha, Tau, DP, Lambda);                                                            
		double S = (h/6)*(fa + 4*fc + fb);                                                                
		return gtFDMaux(a, b, Tol, S, fa, fb, fc, maxRecursionDepth, Alpha, Tau, DP); 
	}
	return Result;
}

/*  double gtFDMaux2(double a, double b, double epsilon, double S, double fa, double fb, double fc, int bottom, double& Alpha, double& Tau, double& DP, double& Lambda) 
{                 
  double c = (a + b)/2, h = b - a;                                                                  
  double d = (a + c)/2, e = (c + b)/2;                                                              
  double fd = AlphaStablePDFInt(d, Alpha,Tau,DP,Lambda), fe = AlphaStablePDFInt(e, Alpha,Tau,DP,Lambda);                                                                      
  double Sleft = (h/12)*(fa + 4*fd + fc);                                                           
  double Sright = (h/12)*(fc + 4*fe + fb);                                                          
  double S2 = Sleft + Sright;                                                                       
  if (bottom <= 0 || fabs(S2 - S) <= 15*epsilon)                                                    
    return S2 + (S2 - S)/15;       
  NumOfLevels++;
  return gtFDMaux2(a, c, epsilon/2, Sleft,  fa, fc, fd, bottom-1, Alpha,Tau,DP,Lambda) +                    
         gtFDMaux2(c, b, epsilon/2, Sright, fc, fb, fe, bottom-1, Alpha,Tau,DP,Lambda); 
} 
  
  double gt_FDMint(double MinAge, double& MaxAge, double& Alpha, double& Tau, double& DP, double& Lambda, double& UZtime)
{
    double Result=0.0;
	if (MinAge > UZtime)
	{
		int maxRecursionDepth=5;
		double a=MinAge-UZtime, b=MaxAge-UZtime, Tol = 1e-03;
		double c = (a + b)/2, h = b - a;                        
		TerminationCode=0, NumOfLevels=0;
		double fa = AlphaStablePDFInt(a,Alpha,Tau,DP,Lambda), fb = AlphaStablePDFInt(b, Alpha, Tau,DP, Lambda), fc = AlphaStablePDFInt(c, Alpha, Tau, DP, Lambda);                                                            
		double S = (h/6)*(fa + 4*fc + fb);                                                                
		return gtFDMaux2(a, b, Tol, S, fa, fb, fc, maxRecursionDepth, Alpha, Tau, DP,Lambda); 
	}
	return Result;
}*/

inline double gt_FDMint(double MinAge, double& MaxAge, double& Alpha, double& Tau, double& DP, double& Lambda, double& UZtime)
{
    double Result=0.0;
	if (MinAge > UZtime)
	{
		//double a=MinAge-UZtime, b=MaxAge-UZtime;
		double k1 = (MaxAge - MinAge)/2, k2 = (MinAge + MaxAge)/2 - UZtime;
		Result+=k1*Coeff3[0]*AlphaStablePDFInt(k1*Roots3[0]+k2,Alpha,Tau,DP,Lambda);
		Result+=k1*Coeff3[1]*AlphaStablePDFInt(k1*Roots3[1]+k2,Alpha,Tau,DP,Lambda);
		Result+=k1*Coeff3[2]*AlphaStablePDFInt(k1*Roots3[2]+k2,Alpha,Tau,DP,Lambda);
	}
	return Result;
}
__declspec(dllexport) double WINAPI FDM(
						FP DateRange[], FP TracerRange[], double Tau, double SampleDate, double Lambda,
						double Alpha, double DP, double UZtime, LPXLOPER12 HeliumThree, LPXLOPER12 InitialTrit, LPXLOPER12 TritInitialTritRatio)
{
	double Result = 1e-012;
	double DR, Cin; // pointers to Date Range and Tracer Input
	double EndDate, FDMprev, CinHe3, FDMnoDecay, TimeIncrement, Lambda2, MaxAge=0.0, Integral, IntegralNoDecay, Lambda0=0.0, UZt=0.0;
	__int32 i,j, nIters, StepInc, StopCriteria;
	nIters = 2000000;
	if ((Tau <= 0) | (DP <= 0) | (SampleDate <= 0) | (DateRange->rows != TracerRange->rows) | (Alpha <= 1) | (Alpha > 2))
	{
		return Result;
	}
	TimeIncrement = 1.0 / 12.0;
    if (Tau >= 100.) // reset time increment and end date for large tau
	{
        TimeIncrement = fabs(DateRange->array[1]-DateRange->array[0]); //MinTimeInc(DateRange);
		if (fabs(DateRange->array[DateRange->rows - 1]-DateRange->array[DateRange->rows - 2]) < TimeIncrement)
		{
			TimeIncrement = fabs(DateRange->array[DateRange->rows - 1]-DateRange->array[DateRange->rows - 2]);
		}
        EndDate = SampleDate - TimeIncrement;
	}
    EndDate = floor (SampleDate) + floor ((SampleDate-floor (SampleDate)) / TimeIncrement) * TimeIncrement - UZtime;
    if (EndDate == SampleDate - UZtime)
	{
        EndDate = EndDate - 1.0/12.0;
	}
    j = DateRange->rows-1;
	StepInc = 1;
	StopCriteria = 0;
	if (DateRange->array[j] < DateRange->array[0])
	{
		StepInc = -1;
		StopCriteria = j;
		j = 0;
	}
	DR = DateRange->array[j];
    if (HeliumThree->val.xbool == TRUE)
	{
        if (Lambda == 0.0)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		for (i = 2; i<= nIters; i++)
		{
            FDMprev = Result;
			MaxAge = i * TimeIncrement;
            while ((EndDate - MaxAge) < DR && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					return Result;
				}
			}
            CinHe3 = Cin * exp(-Lambda2 * UZtime) * (1. - exp(-Lambda2 * (MaxAge - TimeIncrement / 2)));
            Integral = gt_FDMint(MaxAge - TimeIncrement, MaxAge, Alpha, Tau, DP,Lambda0,UZt);
            Result += CinHe3 * Integral;
			if (MaxAge > Tau && Result > 0) //This has to be outer if,then; otherwise Overflow occurs because of division by zero
			{
				if ((Result - FDMprev) / Result < Tol*10)
				{
					return Result;
				}
			}
		}
	}
    if (TritInitialTritRatio->val.xbool == TRUE)
	{
        if (Lambda == 0.0)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
		FDMnoDecay = 0.0;
        for (i = 2; i<= nIters; i++)
		{
            FDMprev = FDMnoDecay;
			MaxAge = i * TimeIncrement;
            while ((EndDate - MaxAge) < DR && j != StopCriteria)
			{
				j = j - StepInc;
				DR = DateRange->array[j];
				Cin = TracerRange->array[j];
				if (Cin == 0 && j == StopCriteria)
				{
					break;
				}
			}
            IntegralNoDecay = gt_FDMint(MaxAge - TimeIncrement, MaxAge, Alpha, Tau, DP,Lambda0,UZt);
            FDMnoDecay = FDMnoDecay + Cin * exp(-Lambda2 * UZtime) * IntegralNoDecay;
            Integral = gt_FDMint(MaxAge - TimeIncrement, MaxAge, Alpha, Tau, DP, Lambda2,UZt);
            Result += Cin * exp(-Lambda2 * UZtime) * Integral;
            if (MaxAge > Tau && FDMnoDecay > 0) //This has to be outer if,then; otherwise Overflow occurs because of division by zero
			{
				if ((FDMnoDecay - FDMprev) / FDMnoDecay < Tol*10)
				{
					break;
				}
			}
		}
        Result = Result / FDMnoDecay;
		return Result;
	}
	if (TritInitialTritRatio->val.xbool != TRUE && HeliumThree->val.xbool != TRUE)
	{
        if (InitialTrit->val.xbool == TRUE)
		{
			Lambda2 = log (2) / 12.32;
		}
		else
		{
			Lambda2 = Lambda;
		}
        for (i = 2; i<= nIters; i++)
		{
            FDMprev = Result;
			MaxAge = i * TimeIncrement;
            while ((EndDate - MaxAge) < DateRange->array[j] && j != StopCriteria)
			{
				j = j - StepInc;
				//DR = DateRange->array[j];
				//Cin = TracerRange->array[j];
				if (TracerRange->array[j] == 0 && j == StopCriteria)
				{
					return Result;
				}
			}
            Result += TracerRange->array[j] * exp(-Lambda2 * UZtime) * gt_FDMint(MaxAge - TimeIncrement, MaxAge, Alpha, Tau, DP, Lambda, UZt);
            if (MaxAge > Tau) //This has to be outer if,then; otherwise Overflow occurs because of division by zero
			{
				if ((Result - FDMprev) / Result < Tol*10)
				{
					return Result;
				}
			}
		}
	}
	return Result;
}
