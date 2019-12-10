
#include "TracerLPM.h"
//#include "unsaturated_zone_gas_tracer-master\unsaturated_zone_tracer_solver.h"


//namespace solver = unsaturated_zone_tracer_solver;

BOOL GetHwnd(HWND * pHwnd)
{
	XLOPER12 x;

	if (Excel12f(xlGetHwnd, &x, 0) == xlretSuccess)
	{
		*pHwnd = (HWND)x.val.w;
		return TRUE;
	}
	return FALSE;
}
///***************************************************************************
// ExcelCursorProc()
//
// Purpose:
//
//      When a modal dialog box is displayed over Microsoft Excel's window, the
//      cursor is a busy cursor over Microsoft Excel's window. This WndProc traps
//      WM_SETCURSORs and changes the cursor back to a normal arrow.
//
// Parameters:
//
//      HWND hWndDlg        Contains the HWND Window
//      UINT message        The message to respond to
//      WPARAM wParam       Arguments passed by Windows
//      LPARAM lParam
//
// Returns: 
//
//      LRESULT             0 if message handled, otherwise the result of the
//                          default WndProc
//
// Comments:
//
// History:  Date       Author        Reason
///***************************************************************************

// Create a place to store Microsoft Excel's WndProc address //
static WNDPROC g_lpfnExcelWndProc = NULL;
#define IDC_PROGBAR 40001

LRESULT CALLBACK ExcelCursorProc(HWND hwnd, 
                                 UINT wMsg, 
                                 WPARAM wParam, 
                                 LPARAM lParam)
{
	//
	// This block checks to see if the message that was passed in is a
	// WM_SETCURSOR message. If so, the cursor is set to an arrow; if not,
	// the default WndProc is called.
	//

	if (wMsg == WM_SETCURSOR)
	{
		SetCursor(LoadCursor(NULL, IDC_ARROW));
		return 0L;
	}
	else
	{
		return CallWindowProc(g_lpfnExcelWndProc, hwnd, wMsg, wParam, lParam);
	}
}
extern void FAR PASCAL HookExcelWindow(HWND hWndExcel)
{
	//
	// This block obtains the address of Microsoft Excel's WndProc through the
	// use of GetWindowLongPtr(). It stores this value in a global that can be
	// used to call the default WndProc and also to restore it. Finally, it
	// replaces this address with the address of ExcelCursorProc using
	// SetWindowLongPtr().
	//

	g_lpfnExcelWndProc = (WNDPROC) GetWindowLongPtr(hWndExcel, GWLP_WNDPROC);
	SetWindowLongPtr(hWndExcel, GWLP_WNDPROC, (LONG_PTR)(FARPROC) ExcelCursorProc);
}

///***************************************************************************
// UnhookExcelWindow()
//
// Purpose:
//
//      This is the function that removes the ExcelCursorProc that was
//      called before Microsoft Excel's main WndProc.
//
// Parameters:
//
//      HANDLE hWndExcel    This is a handle to Microsoft Excel's hWnd
//
// Returns: 
//
// Comments:
//
// History:  Date       Author        Reason
///***************************************************************************

extern void FAR PASCAL UnhookExcelWindow(HWND hWndExcel)
{
	//
	// This function restores Microsoft Excel's default WndProc using
	// SetWindowLongPtr to restore the address that was saved into
	// g_lpfnExcelWndProc by HookExcelWindow(). It then sets g_lpfnExcelWndProc
	// to NULL.
	//

	SetWindowLongPtr(hWndExcel, GWLP_WNDPROC, (LONG_PTR) g_lpfnExcelWndProc);
	g_lpfnExcelWndProc = NULL;
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) {

	static HWND	hwndStart, hwndProgress;
	static HANDLE	hMutex;
	static int	nProgress = 0;
	LOGFONT LogFont;
	HDC         hDC;
    PAINTSTRUCT Ps;
	HFONT hFont;

	switch (message) {
		/*case WM_PAINT:
			hDC = BeginPaint(hwnd, &Ps);
		    
			LogFont.lfStrikeOut = 0;
			LogFont.lfUnderline = 0;
			LogFont.lfHeight = 42;
			LogFont.lfEscapement = 0;
			LogFont.lfWeight = FW_NORMAL;
			
			hFont = CreateFontIndirect(&LogFont);
			SelectObject(hDC, hFont);
			DeleteObject(hFont);

			EndPaint(hwnd, &Ps);
			break;*/
		case WM_CREATE:
		{
			//Initialize progress controls
			INITCOMMONCONTROLSEX iccex;
			iccex.dwSize = sizeof(iccex);
			iccex.dwICC = ICC_PROGRESS_CLASS; // | ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES;
			InitCommonControlsEx(&iccex);

			//Create the mutex
			//hMutex = CreateMutex(NULL,FALSE,NULL);

			//Create the start button
			//hwndStart = CreateWindow(
			//	"BUTTON", "Start", WS_CHILD | WS_VISIBLE, 300, 25, 50, 20, 
			//	hwnd, (HMENU)IDC_START, NULL, NULL);

			HWND hwndLabel = CreateWindowEx(0,L"STATIC",L"Monte Carlo simulation progress", WS_VISIBLE | WS_CHILD | SS_LEFT,
							30,5,350,20,hwnd,NULL,NULL,NULL);
			
			//Create the progress control
			hwndProgress = CreateWindowEx(WS_EX_CLIENTEDGE,
				PROGRESS_CLASS,NULL,WS_CHILD|WS_VISIBLE|PBS_SMOOTH|WS_BORDER, 35,30,250,30,
				hwnd, (HMENU)IDC_PROGBAR, NULL, NULL);

			//Set the progress bar range and initial position
			SendMessage(hwndProgress,(UINT) PBM_SETBARCOLOR,0,(LPARAM)RGB(60,200,100));
			SendMessage(hwndProgress, PBM_SETSTEP, (WPARAM) 1, 0);
			SendMessage(hwndProgress, PBM_SETPOS, 0, 0);
			break;
		}
	
		default:
			return DefWindowProc(hwnd, message, wParam, lParam);
	}

	return 0;
}
HWND CreateProgressBase()
{
	GetHwnd(&g_hWndMain);
	Excel12f(xlEnableXLMsgs, 0, 0);
	HookExcelWindow(g_hWndMain);
	WNDCLASSEX wcex = {0};

	wcex.cbSize		= sizeof(WNDCLASSEX);
	wcex.lpfnWndProc	= WndProc;
	wcex.hInstance		= (HINSTANCE) g_hInst;
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszClassName	= L"MAIN";
	wcex.hIconSm		= LoadIcon(NULL,IDC_ARROW);

	bool Junk2 = RegisterClassEx(&wcex);
	int desktopwidth=GetSystemMetrics(SM_CXSCREEN);
	int desktopheight=GetSystemMetrics(SM_CYSCREEN);
	return CreateWindowEx(WS_EX_OVERLAPPEDWINDOW, L"MAIN", L"TracerLPM", 
            WS_OVERLAPPEDWINDOW, desktopwidth/2-175, desktopheight/2, 350, 130, NULL, NULL, (HINSTANCE) g_hInst, NULL); //g_hWndMain, (HMENU) IDC_PROGBAR, g_hInst2, NULL); // NULL, NULL, g_hInst2, NULL); // WS_CHILD | WS_VISIBLE or WS_OVERLAPPEDWINDOW
}
VectorXd SolveSVD(MatrixXd Xmatrix, VectorXd Bvec)
{
	Matrix2d SVDmat2;
	Matrix3d SVDmat3;
	Matrix4d SVDmat4;
	
	switch (Xmatrix.rows())
	{
	case 1:
		//return Xmatrix.jacobiSvd(ComputeThinU | ComputeThinV).solve(Bvec);
		return VectorXd::Constant(1,1, 1/Xmatrix(0,0))*Bvec;
	case 2:
		SVDmat2 = Xmatrix;
		//e=ATA.jacobiSvd(ComputeThinU | ComputeThinV).solve(ATD);
		return SVDmat2.jacobiSvd(ComputeThinU | ComputeThinV).solve(Bvec);
	case 3:
		SVDmat3 = Xmatrix;
		return SVDmat3.jacobiSvd(ComputeThinU | ComputeThinV).solve(Bvec);
	case 4:
		SVDmat4 = Xmatrix;
		return SVDmat4.jacobiSvd(ComputeThinU | ComputeThinV).solve(Bvec);
	default:
		return Xmatrix.jacobiSvd(ComputeThinU | ComputeThinV).solve(Bvec);
	}
}

MatrixXd InvertMatrix(MatrixXd MatrixToInvert)
{
	Matrix2d Inv2;
	Matrix3d Inv3;
	Matrix4d Inv4;

	switch (MatrixToInvert.rows())
	{
	case 1:
		return MatrixXd::Constant(1,1, 1/MatrixToInvert(0,0));
	case 2:
		Inv2 = MatrixToInvert;
		return Inv2.inverse();
	case 3:
		Inv3 = MatrixToInvert;
		return Inv3.inverse();
	case 4:
		Inv4 = MatrixToInvert;
		return Inv4.inverse();
	default:
		return MatrixToInvert.inverse();
	}
}

__declspec(dllexport) LPXLOPER12 WINAPI SolveNewtonMethod(LPXLOPER12 lxMeasTracerConcs, LPXLOPER12 lxMeasSigmas, 
	LPXLOPER12 lxSampleDates, int ModelNum, LPXLOPER12 lxFitParmIndexes, LPXLOPER12 lxInitModVals, 
	LPXLOPER12 lxLowBounds, LPXLOPER12 lxHiBounds, LPXLOPER12 lxTracers,FP lxdateRange[],LPXLOPER12 lxTracerInputRange, 
	LPXLOPER12 lxLambda, LPXLOPER12 lxuzTime, LPXLOPER12 lxUZtimeCond, LPXLOPER12 lxTracerComp_2, LPXLOPER12 lxDIC_1, LPXLOPER12 lxDIC_2,double Uppm, double THppm, 
	double Porosity, double SedDensity, double He4SolnRate, LPXLOPER12 lxIsMonteCarlo, int iTotalSims, LPXLOPER12 lxIsWriteOut, LPXLOPER12 lxOutFile)
{
	std::clock_t start;
	start = std::clock();

	double HiChiSqr, ChiSqr, PrevChiSqr, Tol, PrevNormV, NormV, NormV_Diff, ScaleFact, Mult;
	double ModelArgs[10] = { 0,0,0,0,0,0,0,0,0,0 }, DoF, test, dLevFactor, ChiSqrDiff, Mean, StdDev;
	int nIters = 0, k, j, i, HiTracer, n, TracerNum = 0, SimCnt = 0, size, MaxIters;
	MatrixXd LJo, InitTracerOutput, Result, ATA, ATD, Jo, JoT, Cov, ATA_Cov;
	VectorXd D, e, sS, V;
	XLOPER12 xMulti[1];
	HWND hwndPB;
	FILE    *MCout = NULL;
	char    Delim[3] = ", ";

#ifdef _DEBUG
	if (fopen_s(&stream, "TracerOutput.txt", "w") == 0)
#endif

		if (lxMeasTracerConcs->val.array.columns*lxMeasTracerConcs->val.array.rows >= 1 && lxdateRange->rows>1 && lxTracerInputRange->val.array.columns*lxTracerInputRange->val.array.rows>1 && lxLambda->xltype == lxuzTime->xltype && lxuzTime->xltype == lxTracers->xltype)
		{

			LPM Solver(ModelNum, lxMeasTracerConcs, lxFitParmIndexes, lxInitModVals, lxTracers, lxdateRange, lxTracerInputRange, lxLambda, lxSampleDates,
				lxuzTime, lxUZtimeCond, lxMeasSigmas, lxHiBounds, lxLowBounds, lxTracerComp_2, Uppm, THppm, Porosity, SedDensity, He4SolnRate, lxIsMonteCarlo, iTotalSims, lxIsWriteOut);

			n = Solver.obj.Model.FitParmIndexes.Val.size();
			TracerNum = Solver.obj.Sample.ActiveVals.sum();
			if (Solver.obj.Model.MonteCarlo.IsMonteCarlo == true)
			{
				size = 4 * n + 7 + (2 * TracerNum);
				if (Solver.obj.Model.MonteCarlo.IsWriteOut == true)
				{
					wchar_t *FileName = deep_copy_wcs(lxOutFile->val.str);
					if (_wfopen_s(&MCout, FileName, L"w") == 0); //Need to add error trap here in case file doesn't open
					else
					{
						char FileName[] = "MonteCarloOut.txt";
						if (fopen_s(&MCout, FileName, "w") == 0);
						else
						{
							printf("The monte carlo output file could not be opened...aborting\n");
							//abort function
						}
					}
					delete FileName;
				}
			}
			else
				size = 2 * n + 6;
			// Create an array of XLOPER12 values.
			XLOPER12 *xOpArray = (XLOPER12 *)malloc(size * sizeof(XLOPER12));
			for (i = 0; i < size; i++)
			{
				xOpArray[i].xltype = xltypeNum;
				xOpArray[i].val.w = -99;
			}
			// Create eigen matrices and vectors
			LJo.resize(TracerNum, 1);
			Jo.resize(TracerNum, n);
			JoT.resize(n, TracerNum);
			D.resize(TracerNum);

			sS.resize(n);
			V.resize(n);

			for (j = 0; j<8; j++)
				ModelArgs[j] = test = Solver.obj.Model.InitModVals(j);
			for (j = 0; j<n; j++)
				V(j) = ModelArgs[Solver.obj.Model.FitParmIndexes.Val[j] - 1];
			if (Solver.obj.Model.FitParmIndexes.isUZtime == true)
			{
				for (i = 0; i<Solver.obj.Tracer.Tracers.size(); i++)
				{
					if (Solver.obj.Tracer.UZtimeCond[i] == 1)
						Solver.obj.Tracer.UZtime(i) = ModelArgs[0];
				}
			}

			InitTracerOutput = Solver.LPM_TracerOutput(ModelArgs[1], ModelArgs[2], ModelArgs[3],
				ModelArgs[4], ModelArgs[5], ModelArgs[6], ModelArgs[7], lxDIC_1->val.num, lxDIC_2->val.num);

#ifdef _DEBUG
			_fprintf_p(stream, "%g", ModelArgs[0]);
			_fprintf_p(stream, "%s", Delim);
			_fprintf_p(stream, "%g", ModelArgs[1]);
			_fprintf_p(stream, "%s", Delim);
			_fprintf_p(stream, "%g", ModelArgs[2]);
			_fprintf_p(stream, "%s", Delim);
			_fprintf_p(stream, "%g", ModelArgs[3]);
			_fprintf_p(stream, "%s", Delim);
			_fprintf_p(stream, "%g", ModelArgs[4]);
			_fprintf_p(stream, "%s", Delim);
			_fprintf_p(stream, "%g", ModelArgs[5]);
			_fprintf_p(stream, "%s", Delim);
			_fprintf_p(stream, "%g", ModelArgs[6]);
			_fprintf_p(stream, "%s", Delim);
			_fprintf_p(stream, "%g", ModelArgs[7]);
			_fprintf_p(stream, "%s\n", "");
#endif
			for (i = 0; i<TracerNum; i++)
			{
				D(i) = (Solver.obj.Sample.MeasTracerConcs(i) - InitTracerOutput(i)) / Solver.obj.Sample.MeasSigmas(i);
				ChiSqr += D(i)*D(i);
#ifdef _DEBUG
				_fprintf_p(stream, "%g", ChiSqr);
				_fprintf_p(stream, "%s", Delim);
				_fprintf_p(stream, "%g\n", InitTracerOutput(i));
#endif
			}
			//eps = std::numeric_limits<double>::epsilon();
			do
			{
				Tol = 0.0001;
				MaxIters = 100;
				NormV = 100;
				PrevNormV = 101;
				NormV_Diff = 100;
				nIters = 0;
				ScaleFact = 0.5;
				dLevFactor = 0.0;
				do
				{
					HiChiSqr = 0;
					nIters++;
					for (j = 0; j < n; j++)
					{
						LJo = Solver.d_dx_LPM_Model(ModelArgs[1], ModelArgs[2], ModelArgs[3],
							ModelArgs[4], ModelArgs[5], ModelArgs[6], ModelArgs[7], lxDIC_1->val.num, lxDIC_2->val.num, Solver.obj.Model.FitParmIndexes.Val[j]);

						for (i = 0; i < TracerNum; i++)
						{
#ifdef _DEBUG
							test = LJo(i);
							_fprintf_p(stream, "%g", test);
							_fprintf_p(stream, "%s", Delim);
#endif
							Jo(i, j) = JoT(j, i) = LJo(i) / Solver.obj.Sample.MeasSigmas(i);
							if (D(i)*D(i) > HiChiSqr)
							{
								HiChiSqr = D(i)*D(i);
								HiTracer = i;
							}
						}
#ifdef _DEBUG
						_fprintf_p(stream, "%s\n", "");
#endif
					}
					ATA = ATA_Cov = JoT * Jo;
					ATD = JoT*D;
					if (dLevFactor > 0)
					{
						for (j = 0; j < n; j++)
						{
							ATA(j, j) += dLevFactor;
						}
					}
					e = ATA.jacobiSvd(ComputeThinU | ComputeThinV).solve(ATD);
					PrevChiSqr = ChiSqr;
					Mult = 1.0;
					for (j = 0; j < n; j++)
					{
						if (Solver.obj.Model.HiBounds((int)Solver.obj.Model.FitParmIndexes.Val[j] - 1) - V(j) < e(j))
						{
							e(j) = ScaleFact*(Solver.obj.Model.HiBounds((int)Solver.obj.Model.FitParmIndexes.Val[j] - 1) - V(j));
						}
						else if (e(j) < Solver.obj.Model.LowBounds((int)Solver.obj.Model.FitParmIndexes.Val[j] - 1) - V(j))
						{
							e(j) = ScaleFact*(Solver.obj.Model.LowBounds((int)Solver.obj.Model.FitParmIndexes.Val[j] - 1) - V(j));
						}
					}
					do
					{
						for (k = 0; k < n; k++)
						{
							ModelArgs[(int)Solver.obj.Model.FitParmIndexes.Val[k] - 1] = V(k) + Mult*e(k);
						}
						if (Solver.obj.Model.FitParmIndexes.isUZtime == true)
						{
							for (j = 0; j < Solver.obj.Tracer.Tracers.size(); j++)
							{
								if (Solver.obj.Tracer.UZtimeCond[j] == 1)
									Solver.obj.Tracer.UZtime(j) = ModelArgs[0];
							}
						}
						ChiSqr = 0;
						InitTracerOutput = Solver.LPM_TracerOutput(ModelArgs[1], ModelArgs[2], ModelArgs[3],
							ModelArgs[4], ModelArgs[5], ModelArgs[6], ModelArgs[7], lxDIC_1->val.num, lxDIC_2->val.num);
#ifdef _DEBUG
						_fprintf_p(stream, "%g", ModelArgs[0]);
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g", ModelArgs[1]);
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g", ModelArgs[2]);
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g", ModelArgs[3]);
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g", ModelArgs[4]);
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g", ModelArgs[5]);
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g", ModelArgs[6]);
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g", ModelArgs[7]);
						_fprintf_p(stream, "%s\n", "");
#endif
						for (int j = 0; j < TracerNum; j++)
						{
							D(j) = (Solver.obj.Sample.MeasTracerConcs(j) - InitTracerOutput(j)) / Solver.obj.Sample.MeasSigmas(j);
							ChiSqr += D(j)*D(j);
#ifdef _DEBUG
							_fprintf_p(stream, "%g", ChiSqr);
							_fprintf_p(stream, "%s", Delim);
							_fprintf_p(stream, "%g\n", InitTracerOutput(j));
#endif
						}
						ChiSqrDiff = ChiSqr - PrevChiSqr;
						Mult *= ScaleFact;
						} while (ChiSqrDiff > 0.0 && Mult > Tol);
						if (Mult < Tol && dLevFactor < 10000)
						{
							if (dLevFactor == 0.0)
								dLevFactor = 1.0;
							else
								dLevFactor *= 10;
							for (k = 0; k < n; k++)
							{
								ModelArgs[(int)Solver.obj.Model.FitParmIndexes.Val[k] - 1] = V(k);
							}
						}
						else
						{
							Mult /= ScaleFact;
							e *= Mult;
							PrevNormV = NormV;
							NormV = e.norm();
							for (i = 0; i < n; i++)
								V(i) = ModelArgs[(int)Solver.obj.Model.FitParmIndexes.Val[i] - 1];
							NormV_Diff = (double)abs(PrevNormV - NormV);
							if (dLevFactor > 0.0  && ChiSqrDiff < 0.0)
								dLevFactor /= 10.0;
						}
				} while (NormV > Tol && NormV_Diff > Tol && nIters < MaxIters);
				//Output begins here
				if (SimCnt == 0)
				{
					ATA = ATA_Cov = JoT * Jo;
					Cov = InvertMatrix(ATA_Cov);
					for (i = 0; i<(2 * n); i += 2)
					{
						xOpArray[i].xltype = xltypeNum;
						xOpArray[i].val.num = V(i / 2);
						xOpArray[i + 1].xltype = xltypeNum;
						xOpArray[i + 1].val.num = sqrt(Cov(i / 2, i / 2));
					}

					xOpArray[2 * n].xltype = xltypeNum;
					xOpArray[2 * n].val.num = ChiSqr;

					DoF = (double)TracerNum - n;  //degrees of freedom
					xOpArray[2 * n + 1].xltype = xltypeNum;
					if (DoF>0 && ChiSqr<100000)
					{
						boost::math::chi_squared_distribution<> ChiSqrDist(DoF);
						xOpArray[2 * n + 1].val.num = 1 - cdf(ChiSqrDist, ChiSqr);
					}
					else
					{
						xOpArray[2 * n + 1].val.num = -99;
					}

					xOpArray[2 * n + 2].xltype = xltypeNum;
					xOpArray[2 * n + 2].val.num = HiTracer;

					xOpArray[2 * n + 3].xltype = xltypeNum;
					xOpArray[2 * n + 3].val.num = HiChiSqr;

					xOpArray[2 * n + 4].xltype = xltypeNum;
					xOpArray[2 * n + 4].val.num = (double)nIters;

					xOpArray[2 * n + 5].xltype = xltypeNum;
					xOpArray[2 * n + 5].val.num = (std::clock() - start) / (double)CLOCKS_PER_SEC;

					if (Solver.obj.Model.MonteCarlo.IsMonteCarlo == true)
					{
						GetHwnd(&g_hWndMain);
						hwndPB = CreateProgressBase();
						SendMessage(GetDlgItem(hwndPB, IDC_PROGBAR), PBM_SETRANGE32, 0, Solver.obj.Model.MonteCarlo.TotalSims);
						ShowWindow(hwndPB, SW_SHOW);
						UpdateWindow(hwndPB);
						if (Solver.obj.Model.MonteCarlo.IsWriteOut == true)
						{
							//Write Header to File
							char s[] = "SimNum, ";
							char t[] = "OptParm, ";
							char x[] = "OptParmErr, ";
							_fprintf_p(MCout, "%s", s);
							for (j = 0; j<n; j++) //Fit parameters
							{
								_fprintf_p(MCout, "%s", t);
								_fprintf_p(MCout, "%s", x);
							}
							char u[] = "ChiSqr, ChiSqrProb, HiTracer, HiChiSqr, Iters, Time, ";
							_fprintf_p(MCout, "%s", u);
							char v[] = "ModConcs, ";
							for (int j = 0; j<TracerNum; j++) //Modeled Tracer Concentrations
							{
								_fprintf_p(MCout, "%s", v);
							}
							char w[] = "SimConcs, ";
							for (int j = 0; j<TracerNum; j++) //Simulated Tracer Concentrations
							{
								_fprintf_p(MCout, "%s", w);
							}
							_fprintf_p(MCout, "%s\n", "");
						}
					}
				}
				else if (Solver.obj.Model.MonteCarlo.IsWriteOut == true)
				{
					//Write Results to File
					for (j = 0; j<n; j++) //Fit parameters
						Solver.obj.Model.MonteCarlo.MonteResults(SimCnt - 1, j) = V(j);
					_fprintf_p(MCout, "%d", SimCnt);
					_fprintf_p(MCout, "%s", Delim);
					Cov = InvertMatrix(ATA);
					for (i = 0; i<(2 * n); i += 2)
					{
						_fprintf_p(MCout, "%g", V(i / 2));
						_fprintf_p(MCout, "%s", Delim);
						_fprintf_p(MCout, "%g", sqrt(Cov(i / 2, i / 2)));
						_fprintf_p(MCout, "%s", Delim);
					}
					_fprintf_p(MCout, "%g", ChiSqr);
					_fprintf_p(MCout, "%s", Delim);

					DoF = (double)TracerNum - n;  //degrees of freedom
					if (DoF>0 && ChiSqr<100000)
					{
						boost::math::chi_squared_distribution<> ChiSqrDist(DoF);
						_fprintf_p(MCout, "%g", 1 - cdf(ChiSqrDist, ChiSqr));
					}
					else
					{
						_fprintf_p(MCout, "%g", (double)-99);
					}
					_fprintf_p(MCout, "%s", Delim);

					_fprintf_p(MCout, "%g", (double)HiTracer);
					_fprintf_p(MCout, "%s", Delim);

					_fprintf_p(MCout, "%g", HiChiSqr);
					_fprintf_p(MCout, "%s", Delim);

					_fprintf_p(MCout, "%g", (double)nIters);
					_fprintf_p(MCout, "%s", Delim);

					_fprintf_p(MCout, "%g", (std::clock() - start) / (double)CLOCKS_PER_SEC);
					_fprintf_p(MCout, "%s", Delim);
					for (int j = 0; j<TracerNum; j++) //Modeled Tracer Concentrations
					{
						Solver.obj.Model.MonteCarlo.MonteResults(SimCnt - 1, j + n) = InitTracerOutput(j);
						_fprintf_p(MCout, "%g", InitTracerOutput(j));
						_fprintf_p(MCout, "%s", Delim);
					}
					for (int j = 0; j<TracerNum; j++) //Simulated Tracer Concentrations
					{
						_fprintf_p(MCout, "%g", Solver.obj.Sample.MeasTracerConcs(j));
						_fprintf_p(MCout, "%s", Delim);
					}
					_fprintf_p(MCout, "%s\n", "");
					}
				else
				{
					for (j = 0; j<n; j++) //Fit parameters
					{
						Solver.obj.Model.MonteCarlo.MonteResults(SimCnt - 1, j) = V(j);
#ifdef _DEBUG
						_fprintf_p(stream, "%g", V(j));
						_fprintf_p(stream, "%s", Delim);
#endif
					}
					for (int j = 0; j<TracerNum; j++) //Modeled Tracer Concentrations
						Solver.obj.Model.MonteCarlo.MonteResults(SimCnt - 1, j + n) = InitTracerOutput(j);
#ifdef _DEBUG
					_fprintf_p(stream, "%s\n", "");
#endif
				}

				//Start Monte Carlo Simulations
				if (Solver.obj.Model.MonteCarlo.IsMonteCarlo == true && SimCnt<Solver.obj.Model.MonteCarlo.TotalSims)
				{
					ChiSqr = 0;
					PrevChiSqr = 0;
					for (int j = 0; j<TracerNum; j++)
					{
						Solver.obj.Sample.MeasTracerConcs(j) = Solver.obj.Model.MonteCarlo.SimulatedConcs(SimCnt, j);
						D(j) = (Solver.obj.Sample.MeasTracerConcs(j) - InitTracerOutput(j)) / Solver.obj.Sample.MeasSigmas(j);
						ChiSqr += D(j)*D(j);
#ifdef _DEBUG
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g\n", Solver.obj.Model.MonteCarlo.SimulatedConcs(SimCnt, j));
#endif
					}
					for (j = 0; j<8; j++)
						ModelArgs[j] = test = Solver.obj.Model.InitModVals(j);
#ifdef _DEBUG
					_fprintf_p(stream, "%s", Delim);
					_fprintf_p(stream, "%g\n", ChiSqr);
#endif
					//SendMessage(hwndPB, PBM_STEPIT, 0, 0);
					SendMessage(GetDlgItem(hwndPB, IDC_PROGBAR), PBM_STEPIT, 0, 0);
					UpdateWindow(hwndPB);
					SimCnt++;
				}
				else if (Solver.obj.Model.MonteCarlo.IsMonteCarlo)
					SimCnt++;
			} while (Solver.obj.Model.MonteCarlo.IsMonteCarlo == true && SimCnt <= Solver.obj.Model.MonteCarlo.TotalSims);
			if (Solver.obj.Model.MonteCarlo.IsMonteCarlo == true)
			{
				xOpArray[2 * n + 6].xltype = xltypeNum;
				xOpArray[2 * n + 6].val.num = (std::clock() - start) / (double)CLOCKS_PER_SEC;
				VectorXd StatVec = VectorXd::Zero(Solver.obj.Model.MonteCarlo.TotalSims);
				j = 2 * (n + TracerNum);
				SimCnt = 0;
				while (j>0)
				{
					StatVec = Solver.obj.Model.MonteCarlo.MonteResults.block(0, SimCnt, Solver.obj.Model.MonteCarlo.TotalSims, 1);
					Mean = StatVec.mean();
					for (int i = 0; i<Solver.obj.Model.MonteCarlo.TotalSims; i++)
						StatVec(i) = pow(Solver.obj.Model.MonteCarlo.MonteResults(i, SimCnt) - Mean, 2);
					StdDev = sqrt(StatVec.sum() / Solver.obj.Model.MonteCarlo.TotalSims);
					xOpArray[4 * n + 7 + (2 * TracerNum) - j].xltype = xltypeNum;
					xOpArray[4 * n + 7 + (2 * TracerNum) - j].val.num = Mean;
					xOpArray[4 * n + 7 + (2 * TracerNum) - j + 1].xltype = xltypeNum;
					xOpArray[4 * n + 7 + (2 * TracerNum) - j + 1].val.num = StdDev;
					j += -2;
					SimCnt++;
				}
				UnhookExcelWindow(g_hWndMain);
				Excel12f(xlDisableXLMsgs, 0, 0);
				DestroyWindow(hwndPB);
				if (Solver.obj.Model.MonteCarlo.IsWriteOut == true)
				{
					fclose(MCout);
				}
			}
			// Create an array of pointers to XLOPER12 values.
			LPXLOPER12 xArray = (LPXLOPER12)malloc(size * sizeof(LPXLOPER12));
			xArray->xltype = xltypeMulti | xlbitDLLFree;
			xArray->val.array.columns = 1;
			xArray->val.array.rows = size;
			xArray->val.array.lparray = xOpArray;

#ifdef _DEBUG
			LPXLOPER12 px;
			_fprintf_p(stream, "%s\n", "");
			for (i = 0; i<2 * n + 6; i++)
			{
				px = xArray->val.array.lparray + i;
				_fprintf_p(stream, "%g\n", px->val.num);
				//_fprintf_p( stream, "%s", Delim );
			}
			fclose(stream);
#endif
			return xArray;
				}
		else
		{
			xMulti[0].val.num = 0;
			xMulti[0].xltype = xltypeNum;
			return (LPXLOPER12)&xMulti[0];
		}
			}


__declspec(dllexport) LPXLOPER12 WINAPI SolveLevenbergMarquardt(LPXLOPER12 lxMeasTracerConcs, LPXLOPER12 lxMeasSigmas,
	LPXLOPER12 lxSampleDates, int ModelNum, LPXLOPER12 lxFitParmIndexes, LPXLOPER12 lxInitModVals,
	LPXLOPER12 lxLowBounds, LPXLOPER12 lxHiBounds, LPXLOPER12 lxTracers, FP lxdateRange[], LPXLOPER12 lxTracerInputRange,
	LPXLOPER12 lxLambda, LPXLOPER12 lxuzTime, LPXLOPER12 lxUZtimeCond, LPXLOPER12 lxTracerComp_2, LPXLOPER12 lxDIC_1, LPXLOPER12 lxDIC_2, double Uppm, double THppm,
	double Porosity, double SedDensity, double He4SolnRate, LPXLOPER12 lxIsMonteCarlo, int iTotalSims, LPXLOPER12 lxIsWriteOut, LPXLOPER12 lxOutFile)
{
	std::clock_t start;
	start = std::clock();

	double HiChiSqr = 0, ChiSqr = 0, PrevChiSqr = 0, Tol, PrevNormV, NormV = 100., NormV_Diff = 100., ScaleFact = 0.4, Mult = 1;
	double ModelArgs[10] = { 0,0,0,0,0,0,0,0,0,0 }, DoF, test, dLevFactor, dLambda, dAlpha, ChiSqrDiff, Mean, StdDev, BestChiSqr, eps;
	int nIters = 0, k, j, i, HiTracer, n, TracerNum = 0, SimCnt = 0, size, MaxIters;
	MatrixXd LJo, InitTracerOutput, Result, ATA, ATD, Jo, JoT, Cov;
	VectorXd D, e, sS, V, BestVals;
	bool IsScale, IsOutBounds;
	XLOPER12 xMulti[1];
	HWND hwndPB;
	FILE    *MCout = NULL;
	char    Delim[3] = ", ";

#ifdef _DEBUG
	if (fopen_s(&stream, "TracerOutput.txt", "w") == 0)
#endif

	if (lxMeasTracerConcs->val.array.columns*lxMeasTracerConcs->val.array.rows >= 1 && lxdateRange->rows>1 && lxTracerInputRange->val.array.columns*lxTracerInputRange->val.array.rows>1 && lxLambda->xltype == lxuzTime->xltype && lxuzTime->xltype == lxTracers->xltype)
	{

		LPM Solver(ModelNum, lxMeasTracerConcs, lxFitParmIndexes, lxInitModVals, lxTracers, lxdateRange, lxTracerInputRange, lxLambda, lxSampleDates,
			lxuzTime, lxUZtimeCond, lxMeasSigmas, lxHiBounds, lxLowBounds, lxTracerComp_2, Uppm, THppm, Porosity, SedDensity, He4SolnRate, lxIsMonteCarlo, iTotalSims, lxIsWriteOut);

		n = Solver.obj.Model.FitParmIndexes.Val.size();
		TracerNum = Solver.obj.Sample.ActiveVals.sum();
		if (Solver.obj.Model.MonteCarlo.IsMonteCarlo == true)
		{
			size = 4 * n + 7 + (2 * TracerNum);
			if (Solver.obj.Model.MonteCarlo.IsWriteOut == true)
			{
				wchar_t *FileName = deep_copy_wcs(lxOutFile->val.str);
				if (_wfopen_s(&MCout, FileName, L"w") == 0); //Need to add error trap here in case file doesn't open
				else
				{
					char FileName[] = "MonteCarloOut.txt";
					if (fopen_s(&MCout, FileName, "w") == 0);
					else
					{
						printf("The monte carlo output file could not be opened...aborting\n");
						//abort function
					}
				}
				delete FileName;
			}
		}
		else
			size = 2 * n + 6;
		// Create an array of XLOPER12 values.
		XLOPER12 *xOpArray = (XLOPER12 *)malloc(size * sizeof(XLOPER12));
		for (i = 0; i < size; i++)
		{
			xOpArray[i].xltype = xltypeNum;
			xOpArray[i].val.w = -99;
		}
		// Create eigen matrices and vectors
		LJo.resize(TracerNum, 1);
		Jo.resize(TracerNum, n);
		JoT.resize(n, TracerNum);
		D.resize(TracerNum);

		sS.resize(n);
		V.resize(n);
		BestVals.resize(n);

		for (j = 0; j<8; j++)
			ModelArgs[j] = test = Solver.obj.Model.InitModVals(j);
		for (j = 0; j<n; j++)
			V(j) = ModelArgs[Solver.obj.Model.FitParmIndexes.Val[j] - 1];
		if (Solver.obj.Model.FitParmIndexes.isUZtime == true)
		{
			for (i = 0; i<Solver.obj.Tracer.Tracers.size(); i++)
			{
				if (Solver.obj.Tracer.UZtimeCond[i] == 1)
					Solver.obj.Tracer.UZtime(i) = ModelArgs[0];
			}
		}

		InitTracerOutput = Solver.LPM_TracerOutput(ModelArgs[1], ModelArgs[2], ModelArgs[3],
			ModelArgs[4], ModelArgs[5], ModelArgs[6], ModelArgs[7], lxDIC_1->val.num, lxDIC_2->val.num);

#ifdef _DEBUG
		_fprintf_p(stream, "%g", ModelArgs[0]);
		_fprintf_p(stream, "%s", Delim);
		_fprintf_p(stream, "%g", ModelArgs[1]);
		_fprintf_p(stream, "%s", Delim);
		_fprintf_p(stream, "%g", ModelArgs[2]);
		_fprintf_p(stream, "%s", Delim);
		_fprintf_p(stream, "%g", ModelArgs[3]);
		_fprintf_p(stream, "%s", Delim);
		_fprintf_p(stream, "%g", ModelArgs[4]);
		_fprintf_p(stream, "%s", Delim);
		_fprintf_p(stream, "%g", ModelArgs[5]);
		_fprintf_p(stream, "%s", Delim);
		_fprintf_p(stream, "%g", ModelArgs[6]);
		_fprintf_p(stream, "%s", Delim);
		_fprintf_p(stream, "%g", ModelArgs[7]);
		_fprintf_p(stream, "%s\n", "");
#endif
		for (i = 0; i<TracerNum; i++)
		{
			D(i) = (Solver.obj.Sample.MeasTracerConcs(i) - InitTracerOutput(i)) / Solver.obj.Sample.MeasSigmas(i);
			ChiSqr += D(i)*D(i);
#ifdef _DEBUG
			_fprintf_p(stream, "%g", ChiSqr);
			_fprintf_p(stream, "%s", Delim);
			_fprintf_p(stream, "%g\n", InitTracerOutput(i));
#endif
		}
		eps = std::numeric_limits<double>::epsilon();
		do
		{
			Tol = 0.0001;
			MaxIters = 50;
			NormV = 100;
			PrevNormV = 101;
			NormV_Diff = 100;
			nIters = 0;
			ScaleFact = 0.5;
			dLevFactor = 1;
			dLambda = 1;
			BestChiSqr = ChiSqr * 2;
			do
			{
				HiChiSqr = 0;
				nIters++;
				for (j = 0; j < n; j++)
				{
					LJo = Solver.d_dx_LPM_Model(ModelArgs[1], ModelArgs[2], ModelArgs[3],
						ModelArgs[4], ModelArgs[5], ModelArgs[6], ModelArgs[7], lxDIC_1->val.num, lxDIC_2->val.num, Solver.obj.Model.FitParmIndexes.Val[j]);

					for (i = 0; i < TracerNum; i++)
					{
#ifdef _DEBUG
						test = LJo(i);
						_fprintf_p(stream, "%g", test);
						_fprintf_p(stream, "%s", Delim);
#endif
						Jo(i, j) = JoT(j, i) = LJo(i) / Solver.obj.Sample.MeasSigmas(i);
						if (D(i)*D(i) > HiChiSqr)
						{
							HiChiSqr = D(i)*D(i);
							HiTracer = i;
						}
					}
#ifdef _DEBUG
					_fprintf_p(stream, "%s\n", "");
#endif
				}
				ATA = JoT * Jo;
				for (j = 0; j < n; j++)
				{
					ATA(j, j) += dLevFactor;
				}
				ATD = JoT*D;
				e = ATA.jacobiSvd(ComputeThinU | ComputeThinV).solve(ATD);
				PrevChiSqr = ChiSqr;
				Mult = 1.0;
				dAlpha = 1.0;
				IsOutBounds = false;
				for (j = 0; j < n; j++)
				{
					if (Solver.obj.Model.HiBounds((int)Solver.obj.Model.FitParmIndexes.Val[j] - 1) - V(j) - Tol < e(j))
					{
						e(j) = (1 - dLambda / 10)*(Solver.obj.Model.HiBounds((int)Solver.obj.Model.FitParmIndexes.Val[j] - 1) - V(j) - Tol);
						dLambda *= ScaleFact;
						IsOutBounds = true;
					}
					else if (e(j) < Solver.obj.Model.LowBounds((int)Solver.obj.Model.FitParmIndexes.Val[j] - 1) - V(j) + Tol)
					{
						e(j) = (1 - dLambda / 10)*(Solver.obj.Model.LowBounds((int)Solver.obj.Model.FitParmIndexes.Val[j] - 1) - V(j) + Tol);
						dLambda *= ScaleFact;
						IsOutBounds = true;
					}
				}
				IsScale = true;
				do
				{
					for (k = 0; k < n; k++)
					{
						ModelArgs[(int)Solver.obj.Model.FitParmIndexes.Val[k] - 1] = V(k) + dAlpha*e(k);
					}
					if (Solver.obj.Model.FitParmIndexes.isUZtime == true)
					{
						for (j = 0; j < Solver.obj.Tracer.Tracers.size(); j++)
						{
							if (Solver.obj.Tracer.UZtimeCond[j] == 1)
								Solver.obj.Tracer.UZtime(j) = ModelArgs[0];
						}
					}
					ChiSqr = 0;
					InitTracerOutput = Solver.LPM_TracerOutput(ModelArgs[1], ModelArgs[2], ModelArgs[3],
						ModelArgs[4], ModelArgs[5], ModelArgs[6], ModelArgs[7], lxDIC_1->val.num, lxDIC_2->val.num);
#ifdef _DEBUG
					_fprintf_p(stream, "%g", ModelArgs[0]);
					_fprintf_p(stream, "%s", Delim);
					_fprintf_p(stream, "%g", ModelArgs[1]);
					_fprintf_p(stream, "%s", Delim);
					_fprintf_p(stream, "%g", ModelArgs[2]);
					_fprintf_p(stream, "%s", Delim);
					_fprintf_p(stream, "%g", ModelArgs[3]);
					_fprintf_p(stream, "%s", Delim);
					_fprintf_p(stream, "%g", ModelArgs[4]);
					_fprintf_p(stream, "%s", Delim);
					_fprintf_p(stream, "%g", ModelArgs[5]);
					_fprintf_p(stream, "%s", Delim);
					_fprintf_p(stream, "%g", ModelArgs[6]);
					_fprintf_p(stream, "%s", Delim);
					_fprintf_p(stream, "%g", ModelArgs[7]);
					_fprintf_p(stream, "%s\n", "");
#endif
					for (int j = 0; j < TracerNum; j++)
					{
						D(j) = (Solver.obj.Sample.MeasTracerConcs(j) - InitTracerOutput(j)) / Solver.obj.Sample.MeasSigmas(j);
						ChiSqr += D(j)*D(j);
#ifdef _DEBUG
						_fprintf_p(stream, "%g", ChiSqr);
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g\n", InitTracerOutput(j));
#endif
					}
					if (IsScale)
					{
						dAlpha = ScaleFact*Mult;
						Mult *= ScaleFact;
						IsScale = false;
					}
					else
					{
						dAlpha = 1.0 - ScaleFact*Mult;
						IsScale = true;
					}
					ChiSqrDiff = ChiSqr - PrevChiSqr;
				} while (ChiSqrDiff > 0.0 && Mult > Tol);
				if (ChiSqrDiff > 0.0 || ChiSqr < BestChiSqr || IsOutBounds == false)
				{
					Mult /= ScaleFact;
					if (IsScale || Mult < 1.0)
					{
						if (IsScale)
							dAlpha = ScaleFact*Mult;
						else
							dAlpha = 1.0 - ScaleFact*Mult;
					}
					else
					{
						dAlpha = Mult;
					}
					e *= dAlpha;
					PrevNormV = NormV;
					NormV = e.norm();
					for (i = 0; i < n; i++)
						V(i) = ModelArgs[(int)Solver.obj.Model.FitParmIndexes.Val[i] - 1];
					NormV_Diff = (double)abs(PrevNormV - NormV);
					if (ChiSqr < BestChiSqr)
					{
						BestChiSqr = ChiSqr;
						BestVals = V;
					}
					if (dLambda < 1.0)
						dLambda /= ScaleFact;
					dLevFactor /= 10;
				}
				else if (ChiSqrDiff > 0.0 && nIters < MaxIters - 1 && IsOutBounds)
				{
					//try backtracking
					for (j = 0; j < n; j++)
					{
						if (e(j) > Solver.obj.Model.HiBounds((int)Solver.obj.Model.FitParmIndexes.Val[j] - 1) - V(j) - Tol)
						{
							//increasing
							e(j) = 0.1*(Solver.obj.Model.LowBounds((int)Solver.obj.Model.FitParmIndexes.Val[j] - 1) - V(j));
						}
						else if (e(j) < Solver.obj.Model.LowBounds((int)Solver.obj.Model.FitParmIndexes.Val[j] - 1) - V(j) + Tol)
						{
							e(j) = 0.1*(Solver.obj.Model.HiBounds((int)Solver.obj.Model.FitParmIndexes.Val[j] - 1) - V(j));
						}
					}
					V += e;
					for (k = 0; k < n; k++)
					{
						ModelArgs[(int)Solver.obj.Model.FitParmIndexes.Val[k] - 1] = V(k);
					}
					if (Solver.obj.Model.FitParmIndexes.isUZtime == true)
					{
						for (j = 0; j < Solver.obj.Tracer.Tracers.size(); j++)
						{
							if (Solver.obj.Tracer.UZtimeCond[j] == 1)
								Solver.obj.Tracer.UZtime(j) = ModelArgs[0];
						}
					}
					ChiSqr = 0;
					InitTracerOutput = Solver.LPM_TracerOutput(ModelArgs[1], ModelArgs[2], ModelArgs[3],
						ModelArgs[4], ModelArgs[5], ModelArgs[6], ModelArgs[7], lxDIC_1->val.num, lxDIC_2->val.num);
					for (int j = 0; j < TracerNum; j++)
					{
						D(j) = (Solver.obj.Sample.MeasTracerConcs(j) - InitTracerOutput(j)) / Solver.obj.Sample.MeasSigmas(j);
						ChiSqr += D(j)*D(j);
#ifdef _DEBUG
						_fprintf_p(stream, "%g", ChiSqr);
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g\n", InitTracerOutput(j));
#endif
					}
					dLevFactor *= 10;
					dLambda = 1;
				}
				else
				{
					V = BestVals;
					for (k = 0; k < n; k++)
					{
						ModelArgs[(int)Solver.obj.Model.FitParmIndexes.Val[k] - 1] = V(k);
					}
					if (Solver.obj.Model.FitParmIndexes.isUZtime == true)
					{
						for (j = 0; j < Solver.obj.Tracer.Tracers.size(); j++)
						{
							if (Solver.obj.Tracer.UZtimeCond[j] == 1)
								Solver.obj.Tracer.UZtime(j) = ModelArgs[0];
						}
					}
					ChiSqr = 0;
					InitTracerOutput = Solver.LPM_TracerOutput(ModelArgs[1], ModelArgs[2], ModelArgs[3],
						ModelArgs[4], ModelArgs[5], ModelArgs[6], ModelArgs[7], lxDIC_1->val.num, lxDIC_2->val.num);
					for (int j = 0; j < TracerNum; j++)
					{
						D(j) = (Solver.obj.Sample.MeasTracerConcs(j) - InitTracerOutput(j)) / Solver.obj.Sample.MeasSigmas(j);
						ChiSqr += D(j)*D(j);
#ifdef _DEBUG
						_fprintf_p(stream, "%g", ChiSqr);
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g\n", InitTracerOutput(j));
#endif
					}
				}
		} while (NormV > Tol && NormV_Diff > Tol && nIters < MaxIters);
//Output begins here
		if (SimCnt == 0)
		{
			Cov = InvertMatrix(ATA);
			for (i = 0; i<(2 * n); i += 2)
			{
				xOpArray[i].xltype = xltypeNum;
				xOpArray[i].val.num = V(i / 2);
				xOpArray[i + 1].xltype = xltypeNum;
				xOpArray[i + 1].val.num = sqrt(Cov(i / 2, i / 2));
			}

			xOpArray[2 * n].xltype = xltypeNum;
			xOpArray[2 * n].val.num = ChiSqr;

			DoF = (double)TracerNum - n;  //degrees of freedom
											//test = pdf(ChiSqrDist,ChiSqr);
											//test = 1-cdf(ChiSqrDist,ChiSqr);
			xOpArray[2 * n + 1].xltype = xltypeNum;
			if (DoF>0 && ChiSqr<100000)
			{
				boost::math::chi_squared_distribution<> ChiSqrDist(DoF);
				xOpArray[2 * n + 1].val.num = 1 - cdf(ChiSqrDist, ChiSqr);
			}
			else
			{
				xOpArray[2 * n + 1].val.num = -99;
			}

			xOpArray[2 * n + 2].xltype = xltypeNum;
			xOpArray[2 * n + 2].val.num = HiTracer;

			xOpArray[2 * n + 3].xltype = xltypeNum;
			xOpArray[2 * n + 3].val.num = HiChiSqr;

			xOpArray[2 * n + 4].xltype = xltypeNum;
			xOpArray[2 * n + 4].val.num = (double)nIters;

			xOpArray[2 * n + 5].xltype = xltypeNum;
			xOpArray[2 * n + 5].val.num = (std::clock() - start) / (double)CLOCKS_PER_SEC;

			if (Solver.obj.Model.MonteCarlo.IsMonteCarlo == true)
			{
				GetHwnd(&g_hWndMain);
				hwndPB = CreateProgressBase();
				SendMessage(GetDlgItem(hwndPB, IDC_PROGBAR), PBM_SETRANGE32, 0, Solver.obj.Model.MonteCarlo.TotalSims);
				ShowWindow(hwndPB, SW_SHOW);
				UpdateWindow(hwndPB);
				if (Solver.obj.Model.MonteCarlo.IsWriteOut == true)
				{
					//Write Header to File
					char s[] = "SimNum, ";
					char t[] = "OptParm, ";
					char x[] = "OptParmErr, ";
					_fprintf_p(MCout, "%s", s);
					for (j = 0; j<n; j++) //Fit parameters
					{
						_fprintf_p(MCout, "%s", t);
						_fprintf_p(MCout, "%s", x);
					}
					char u[] = "ChiSqr, ChiSqrProb, HiTracer, HiChiSqr, Iters, Time, ";
					_fprintf_p(MCout, "%s", u);
					char v[] = "ModConcs, ";
					for (int j = 0; j<TracerNum; j++) //Modeled Tracer Concentrations
					{
						_fprintf_p(MCout, "%s", v);
					}
					char w[] = "SimConcs, ";
					for (int j = 0; j<TracerNum; j++) //Simulated Tracer Concentrations
					{
						_fprintf_p(MCout, "%s", w);
					}
					_fprintf_p(MCout, "%s\n", "");
				}
			}
		}
		else if (Solver.obj.Model.MonteCarlo.IsWriteOut == true)
		{
			//Write Results to File
			for (j = 0; j<n; j++) //Fit parameters
				Solver.obj.Model.MonteCarlo.MonteResults(SimCnt - 1, j) = V(j);
			_fprintf_p(MCout, "%d", SimCnt);
			_fprintf_p(MCout, "%s", Delim);
			Cov = InvertMatrix(ATA);
			for (i = 0; i<(2 * n); i += 2)
			{
				_fprintf_p(MCout, "%g", V(i / 2));
				_fprintf_p(MCout, "%s", Delim);
				_fprintf_p(MCout, "%g", sqrt(Cov(i / 2, i / 2)));
				_fprintf_p(MCout, "%s", Delim);
			}
			_fprintf_p(MCout, "%g", ChiSqr);
			_fprintf_p(MCout, "%s", Delim);

			DoF = (double)TracerNum - n;  //degrees of freedom
			if (DoF>0 && ChiSqr<100000)
			{
				boost::math::chi_squared_distribution<> ChiSqrDist(DoF);
				_fprintf_p(MCout, "%g", 1 - cdf(ChiSqrDist, ChiSqr));
			}
			else
			{
				_fprintf_p(MCout, "%g", (double)-99);
			}
			_fprintf_p(MCout, "%s", Delim);

			_fprintf_p(MCout, "%g", (double)HiTracer);
			_fprintf_p(MCout, "%s", Delim);

			_fprintf_p(MCout, "%g", HiChiSqr);
			_fprintf_p(MCout, "%s", Delim);

			_fprintf_p(MCout, "%g", (double)nIters);
			_fprintf_p(MCout, "%s", Delim);

			_fprintf_p(MCout, "%g", (std::clock() - start) / (double)CLOCKS_PER_SEC);
			_fprintf_p(MCout, "%s", Delim);
			for (int j = 0; j<TracerNum; j++) //Modeled Tracer Concentrations
			{
				Solver.obj.Model.MonteCarlo.MonteResults(SimCnt - 1, j + n) = InitTracerOutput(j);
				_fprintf_p(MCout, "%g", InitTracerOutput(j));
				_fprintf_p(MCout, "%s", Delim);
			}
			for (int j = 0; j<TracerNum; j++) //Simulated Tracer Concentrations
			{
				_fprintf_p(MCout, "%g", Solver.obj.Sample.MeasTracerConcs(j));
				_fprintf_p(MCout, "%s", Delim);
			}
			_fprintf_p(MCout, "%s\n", "");
			}
		else
		{
			for (j = 0; j<n; j++) //Fit parameters
			{
				Solver.obj.Model.MonteCarlo.MonteResults(SimCnt - 1, j) = V(j);
#ifdef _DEBUG
				_fprintf_p(stream, "%g", V(j));
				_fprintf_p(stream, "%s", Delim);
#endif
			}
			for (int j = 0; j<TracerNum; j++) //Modeled Tracer Concentrations
				Solver.obj.Model.MonteCarlo.MonteResults(SimCnt - 1, j + n) = InitTracerOutput(j);
#ifdef _DEBUG
			_fprintf_p(stream, "%s\n", "");
#endif
		}

		//Start Monte Carlo Simulations
		if (Solver.obj.Model.MonteCarlo.IsMonteCarlo == true && SimCnt<Solver.obj.Model.MonteCarlo.TotalSims)
		{
			ChiSqr = 0;
			PrevChiSqr = 0;
			for (int j = 0; j<TracerNum; j++)
			{
				Solver.obj.Sample.MeasTracerConcs(j) = Solver.obj.Model.MonteCarlo.SimulatedConcs(SimCnt, j);
				D(j) = (Solver.obj.Sample.MeasTracerConcs(j) - InitTracerOutput(j)) / Solver.obj.Sample.MeasSigmas(j);
				ChiSqr += D(j)*D(j);
#ifdef _DEBUG
				_fprintf_p(stream, "%s", Delim);
				_fprintf_p(stream, "%g\n", Solver.obj.Model.MonteCarlo.SimulatedConcs(SimCnt, j));
#endif
			}
			for (j = 0; j<8; j++)
				ModelArgs[j] = test = Solver.obj.Model.InitModVals(j);
#ifdef _DEBUG
			_fprintf_p(stream, "%s", Delim);
			_fprintf_p(stream, "%g\n", ChiSqr);
#endif
			//SendMessage(hwndPB, PBM_STEPIT, 0, 0);
			SendMessage(GetDlgItem(hwndPB, IDC_PROGBAR), PBM_STEPIT, 0, 0);
			UpdateWindow(hwndPB);
			SimCnt++;
		}
		else if (Solver.obj.Model.MonteCarlo.IsMonteCarlo)
		SimCnt++;
		}while (Solver.obj.Model.MonteCarlo.IsMonteCarlo == true && SimCnt <= Solver.obj.Model.MonteCarlo.TotalSims);
		if (Solver.obj.Model.MonteCarlo.IsMonteCarlo == true)
		{
			xOpArray[2 * n + 6].xltype = xltypeNum;
			xOpArray[2 * n + 6].val.num = (std::clock() - start) / (double)CLOCKS_PER_SEC;
			VectorXd StatVec = VectorXd::Zero(Solver.obj.Model.MonteCarlo.TotalSims);
			j = 2 * (n + TracerNum);
			SimCnt = 0;
			while (j>0)
			{
				StatVec = Solver.obj.Model.MonteCarlo.MonteResults.block(0, SimCnt, Solver.obj.Model.MonteCarlo.TotalSims, 1);
				Mean = StatVec.mean();
				for (int i = 0; i<Solver.obj.Model.MonteCarlo.TotalSims; i++)
					StatVec(i) = pow(Solver.obj.Model.MonteCarlo.MonteResults(i, SimCnt) - Mean, 2);
				StdDev = sqrt(StatVec.sum() / Solver.obj.Model.MonteCarlo.TotalSims);
				xOpArray[4 * n + 7 + (2 * TracerNum) - j].xltype = xltypeNum;
				xOpArray[4 * n + 7 + (2 * TracerNum) - j].val.num = Mean;
				xOpArray[4 * n + 7 + (2 * TracerNum) - j + 1].xltype = xltypeNum;
				xOpArray[4 * n + 7 + (2 * TracerNum) - j + 1].val.num = StdDev;
				j += -2;
				SimCnt++;
			}
			UnhookExcelWindow(g_hWndMain);
			Excel12f(xlDisableXLMsgs, 0, 0);
			DestroyWindow(hwndPB);
			if (Solver.obj.Model.MonteCarlo.IsWriteOut == true)
			{
				fclose(MCout);
			}
		}
		// Create an array of pointers to XLOPER12 values.
		LPXLOPER12 xArray = (LPXLOPER12)malloc(size * sizeof(LPXLOPER12));
		xArray->xltype = xltypeMulti | xlbitDLLFree;
		xArray->val.array.columns = 1;
		xArray->val.array.rows = size;
		xArray->val.array.lparray = xOpArray;

		#ifdef _DEBUG
		LPXLOPER12 px;
		_fprintf_p(stream, "%s\n", "");
		for (i = 0; i<2 * n + 6; i++)
		{
			px = xArray->val.array.lparray + i;
			_fprintf_p(stream, "%g\n", px->val.num);
			//_fprintf_p( stream, "%s", Delim );
		}
		fclose(stream);
		#endif
		return xArray;
	}
	else
	{
		xMulti[0].val.num = 0;
		xMulti[0].xltype = xltypeNum;
		return (LPXLOPER12)&xMulti[0];
	}
}


__declspec(dllexport) LPXLOPER12 WINAPI SolveGNLM(LPXLOPER12 lxMeasTracerConcs, LPXLOPER12 lxMeasSigmas,
	LPXLOPER12 lxSampleDates, int ModelNum, LPXLOPER12 lxFitParmIndexes, LPXLOPER12 lxInitModVals,
	LPXLOPER12 lxLowBounds, LPXLOPER12 lxHiBounds, LPXLOPER12 lxTracers, FP lxdateRange[], LPXLOPER12 lxTracerInputRange,
	LPXLOPER12 lxLambda, LPXLOPER12 lxuzTime, LPXLOPER12 lxUZtimeCond, LPXLOPER12 lxTracerComp_2, LPXLOPER12 lxDIC_1, LPXLOPER12 lxDIC_2, double Uppm, double THppm,
	double Porosity, double SedDensity, double He4SolnRate, LPXLOPER12 lxIsMonteCarlo, int iTotalSims, LPXLOPER12 lxIsWriteOut, LPXLOPER12 lxOutFile)
{
	std::clock_t start;
	start = std::clock();

	double HiChiSqr, ChiSqr, PrevChiSqr, Tol, PrevNormV, NormV, NormV_Diff, ScaleFact, Mult;
	double ModelArgs[10] = { 0,0,0,0,0,0,0,0,0,0 }, DoF, test, dLevFactor, dAlpha, dLastAlpha, ChiSqrDiff, Mean, StdDev;
	int nIters = 0, k, j, i, HiTracer, n, TracerNum = 0, SimCnt = 0, size, MaxIters;
	bool IsScale;
	MatrixXd LJo, InitTracerOutput, Result, ATA, ATD, Jo, JoT, Cov, ATA_Cov;
	VectorXd D, e, sS, V;
	XLOPER12 xMulti[1];
	HWND hwndPB;
	FILE    *MCout = NULL;
	char    Delim[3] = ", ";

#ifdef _DEBUG
	if (fopen_s(&stream, "TracerOutput.txt", "w") == 0)
#endif

		if (lxMeasTracerConcs->val.array.columns*lxMeasTracerConcs->val.array.rows >= 1 && lxdateRange->rows>1 && lxTracerInputRange->val.array.columns*lxTracerInputRange->val.array.rows>1 && lxLambda->xltype == lxuzTime->xltype && lxuzTime->xltype == lxTracers->xltype)
		{

			LPM Solver(ModelNum, lxMeasTracerConcs, lxFitParmIndexes, lxInitModVals, lxTracers, lxdateRange, lxTracerInputRange, lxLambda, lxSampleDates,
				lxuzTime, lxUZtimeCond, lxMeasSigmas, lxHiBounds, lxLowBounds, lxTracerComp_2, Uppm, THppm, Porosity, SedDensity, He4SolnRate, lxIsMonteCarlo, iTotalSims, lxIsWriteOut);

			n = Solver.obj.Model.FitParmIndexes.Val.size();
			TracerNum = Solver.obj.Sample.ActiveVals.sum();
			if (Solver.obj.Model.MonteCarlo.IsMonteCarlo == true)
			{
				size = 4 * n + 7 + (2 * TracerNum);
				if (Solver.obj.Model.MonteCarlo.IsWriteOut == true)
				{
					wchar_t *FileName = deep_copy_wcs(lxOutFile->val.str);
					if (_wfopen_s(&MCout, FileName, L"w") == 0); //Need to add error trap here in case file doesn't open
					else
					{
						char FileName[] = "MonteCarloOut.txt";
						if (fopen_s(&MCout, FileName, "w") == 0);
						else
						{
							printf("The monte carlo output file could not be opened...aborting\n");
							//abort function
						}
					}
					delete FileName;
				}
			}
			else
				size = 2 * n + 6;
			// Create an array of XLOPER12 values.
			XLOPER12 *xOpArray = (XLOPER12 *)malloc(size * sizeof(XLOPER12));
			for (i = 0; i < size; i++)
			{
				xOpArray[i].xltype = xltypeNum;
				xOpArray[i].val.w = -99;
			}
			// Create eigen matrices and vectors
			LJo.resize(TracerNum, 1);
			Jo.resize(TracerNum, n);
			JoT.resize(n, TracerNum);
			D.resize(TracerNum);

			sS.resize(n);
			V.resize(n);

			for (j = 0; j<8; j++)
				ModelArgs[j] = test = Solver.obj.Model.InitModVals(j);
			for (j = 0; j<n; j++)
				V(j) = ModelArgs[Solver.obj.Model.FitParmIndexes.Val[j] - 1];
			if (Solver.obj.Model.FitParmIndexes.isUZtime == true)
			{
				for (i = 0; i<Solver.obj.Tracer.Tracers.size(); i++)
				{
					if (Solver.obj.Tracer.UZtimeCond[i] == 1)
						Solver.obj.Tracer.UZtime(i) = ModelArgs[0];
				}
			}

			InitTracerOutput = Solver.LPM_TracerOutput(ModelArgs[1], ModelArgs[2], ModelArgs[3],
				ModelArgs[4], ModelArgs[5], ModelArgs[6], ModelArgs[7], lxDIC_1->val.num, lxDIC_2->val.num);

#ifdef _DEBUG
			_fprintf_p(stream, "%g", ModelArgs[0]);
			_fprintf_p(stream, "%s", Delim);
			_fprintf_p(stream, "%g", ModelArgs[1]);
			_fprintf_p(stream, "%s", Delim);
			_fprintf_p(stream, "%g", ModelArgs[2]);
			_fprintf_p(stream, "%s", Delim);
			_fprintf_p(stream, "%g", ModelArgs[3]);
			_fprintf_p(stream, "%s", Delim);
			_fprintf_p(stream, "%g", ModelArgs[4]);
			_fprintf_p(stream, "%s", Delim);
			_fprintf_p(stream, "%g", ModelArgs[5]);
			_fprintf_p(stream, "%s", Delim);
			_fprintf_p(stream, "%g", ModelArgs[6]);
			_fprintf_p(stream, "%s", Delim);
			_fprintf_p(stream, "%g", ModelArgs[7]);
			_fprintf_p(stream, "%s\n", "");
#endif
			for (i = 0; i<TracerNum; i++)
			{
				D(i) = (Solver.obj.Sample.MeasTracerConcs(i) - InitTracerOutput(i)) / Solver.obj.Sample.MeasSigmas(i);
				ChiSqr += D(i)*D(i);
#ifdef _DEBUG
				_fprintf_p(stream, "%g", ChiSqr);
				_fprintf_p(stream, "%s", Delim);
				_fprintf_p(stream, "%g\n", InitTracerOutput(i));
#endif
			}
			//eps = std::numeric_limits<double>::epsilon();
			do
			{
				Tol = 0.0001;
				MaxIters = 50;
				NormV = 100;
				PrevNormV = 101;
				NormV_Diff = 100;
				nIters = 0;
				ScaleFact = 0.5;
				dLevFactor = 0.0;
				do
				{
					HiChiSqr = 0;
					nIters++;
					for (j = 0; j < n; j++)
					{
						LJo = Solver.d_dx_LPM_Model(ModelArgs[1], ModelArgs[2], ModelArgs[3],
							ModelArgs[4], ModelArgs[5], ModelArgs[6], ModelArgs[7], lxDIC_1->val.num, lxDIC_2->val.num, Solver.obj.Model.FitParmIndexes.Val[j]);

						for (i = 0; i < TracerNum; i++)
						{
#ifdef _DEBUG
							test = LJo(i);
							_fprintf_p(stream, "%g", test);
							_fprintf_p(stream, "%s", Delim);
#endif
							Jo(i, j) = JoT(j, i) = LJo(i) / Solver.obj.Sample.MeasSigmas(i);
							if (D(i)*D(i) > HiChiSqr)
							{
								HiChiSqr = D(i)*D(i);
								HiTracer = i;
							}
						}
#ifdef _DEBUG
						_fprintf_p(stream, "%s\n", "");
#endif
					}
					ATA = ATA_Cov = JoT * Jo;
					ATD = JoT*D;
					if (dLevFactor > 0)
					{
						for (j = 0; j < n; j++)
						{
							ATA(j, j) += dLevFactor;
						}
					}
					e = ATA.jacobiSvd(ComputeThinU | ComputeThinV).solve(ATD);
					PrevChiSqr = ChiSqr;
					Mult = 1.0;
					dAlpha = 1.0;
					for (j = 0; j < n; j++)
					{
						if (Solver.obj.Model.HiBounds((int)Solver.obj.Model.FitParmIndexes.Val[j] - 1) - V(j) < e(j))
						{
							e(j) = ScaleFact*(Solver.obj.Model.HiBounds((int)Solver.obj.Model.FitParmIndexes.Val[j] - 1) - V(j));
						}
						else if (e(j) < Solver.obj.Model.LowBounds((int)Solver.obj.Model.FitParmIndexes.Val[j] - 1) - V(j))
						{
							e(j) = ScaleFact*(Solver.obj.Model.LowBounds((int)Solver.obj.Model.FitParmIndexes.Val[j] - 1) - V(j));
						}
					}
					IsScale = true;
					do
					{
						dLastAlpha = dAlpha;
						for (k = 0; k < n; k++)
						{
							ModelArgs[(int)Solver.obj.Model.FitParmIndexes.Val[k] - 1] = V(k) + dLastAlpha*e(k);
						}
						if (Solver.obj.Model.FitParmIndexes.isUZtime == true)
						{
							for (j = 0; j < Solver.obj.Tracer.Tracers.size(); j++)
							{
								if (Solver.obj.Tracer.UZtimeCond[j] == 1)
									Solver.obj.Tracer.UZtime(j) = ModelArgs[0];
							}
						}
						ChiSqr = 0;
						InitTracerOutput = Solver.LPM_TracerOutput(ModelArgs[1], ModelArgs[2], ModelArgs[3],
							ModelArgs[4], ModelArgs[5], ModelArgs[6], ModelArgs[7], lxDIC_1->val.num, lxDIC_2->val.num);
#ifdef _DEBUG
						_fprintf_p(stream, "%g", ModelArgs[0]);
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g", ModelArgs[1]);
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g", ModelArgs[2]);
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g", ModelArgs[3]);
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g", ModelArgs[4]);
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g", ModelArgs[5]);
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g", ModelArgs[6]);
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g", ModelArgs[7]);
						_fprintf_p(stream, "%s\n", "");
#endif
						for (int j = 0; j < TracerNum; j++)
						{
							D(j) = (Solver.obj.Sample.MeasTracerConcs(j) - InitTracerOutput(j)) / Solver.obj.Sample.MeasSigmas(j);
							ChiSqr += D(j)*D(j);
#ifdef _DEBUG
							_fprintf_p(stream, "%g", ChiSqr);
							_fprintf_p(stream, "%s", Delim);
							_fprintf_p(stream, "%g\n", InitTracerOutput(j));
#endif
						}
						if (IsScale)
						{
							Mult *= ScaleFact; 
							dAlpha = Mult;
							IsScale = false;
						}
						else
						{
							dAlpha = 1.0 - ScaleFact*Mult;
							IsScale = true;
						}
						ChiSqrDiff = ChiSqr - PrevChiSqr;
					} while (ChiSqrDiff > 0.0 && Mult > Tol);
					if (Mult < Tol && dLevFactor < 10000)
					{
						if (dLevFactor == 0.0)
							dLevFactor = 0.01;
						else
							dLevFactor *= 10;
						for (k = 0; k < n; k++)
						{
							V(k) = ModelArgs[(int)Solver.obj.Model.FitParmIndexes.Val[k] - 1];
						}
					}
					else
					{
						e *= dLastAlpha;
						PrevNormV = NormV;
						NormV = e.norm();
						for (i = 0; i < n; i++)
							V(i) = ModelArgs[(int)Solver.obj.Model.FitParmIndexes.Val[i] - 1];
						NormV_Diff = (double)abs(PrevNormV - NormV);
						if (dLevFactor > 0.0  && ChiSqrDiff < 0.0)
							dLevFactor /= 10.0;
					}
				} while (NormV > Tol && NormV_Diff > Tol && nIters < MaxIters);
				//Output begins here
				if (SimCnt == 0)
				{
					ATA = ATA_Cov = JoT * Jo;
					Cov = InvertMatrix(ATA);
					for (i = 0; i<(2 * n); i += 2)
					{
						xOpArray[i].xltype = xltypeNum;
						xOpArray[i].val.num = V(i / 2);
						xOpArray[i + 1].xltype = xltypeNum;
						xOpArray[i + 1].val.num = sqrt(Cov(i / 2, i / 2));
					}

					xOpArray[2 * n].xltype = xltypeNum;
					xOpArray[2 * n].val.num = ChiSqr;

					DoF = (double)TracerNum - n;  //degrees of freedom
												  //test = pdf(ChiSqrDist,ChiSqr);
												  //test = 1-cdf(ChiSqrDist,ChiSqr);
					xOpArray[2 * n + 1].xltype = xltypeNum;
					if (DoF>0 && ChiSqr<100000)
					{
						boost::math::chi_squared_distribution<> ChiSqrDist(DoF);
						xOpArray[2 * n + 1].val.num = 1 - cdf(ChiSqrDist, ChiSqr);
					}
					else
					{
						xOpArray[2 * n + 1].val.num = -99;
					}

					xOpArray[2 * n + 2].xltype = xltypeNum;
					xOpArray[2 * n + 2].val.num = HiTracer;

					xOpArray[2 * n + 3].xltype = xltypeNum;
					xOpArray[2 * n + 3].val.num = HiChiSqr;

					xOpArray[2 * n + 4].xltype = xltypeNum;
					xOpArray[2 * n + 4].val.num = (double)nIters;

					xOpArray[2 * n + 5].xltype = xltypeNum;
					xOpArray[2 * n + 5].val.num = (std::clock() - start) / (double)CLOCKS_PER_SEC;

					if (Solver.obj.Model.MonteCarlo.IsMonteCarlo == true)
					{
						GetHwnd(&g_hWndMain);
						hwndPB = CreateProgressBase();
						SendMessage(GetDlgItem(hwndPB, IDC_PROGBAR), PBM_SETRANGE32, 0, Solver.obj.Model.MonteCarlo.TotalSims);
						ShowWindow(hwndPB, SW_SHOW);
						UpdateWindow(hwndPB);
						if (Solver.obj.Model.MonteCarlo.IsWriteOut == true)
						{
							//Write Header to File
							char s[] = "SimNum, ";
							char t[] = "OptParm, ";
							char x[] = "OptParmErr, ";
							_fprintf_p(MCout, "%s", s);
							for (j = 0; j<n; j++) //Fit parameters
							{
								_fprintf_p(MCout, "%s", t);
								_fprintf_p(MCout, "%s", x);
							}
							char u[] = "ChiSqr, ChiSqrProb, HiTracer, HiChiSqr, Iters, Time, ";
							_fprintf_p(MCout, "%s", u);
							char v[] = "ModConcs, ";
							for (int j = 0; j<TracerNum; j++) //Modeled Tracer Concentrations
							{
								_fprintf_p(MCout, "%s", v);
							}
							char w[] = "SimConcs, ";
							for (int j = 0; j<TracerNum; j++) //Simulated Tracer Concentrations
							{
								_fprintf_p(MCout, "%s", w);
							}
							_fprintf_p(MCout, "%s\n", "");
						}
					}
				}
				else if (Solver.obj.Model.MonteCarlo.IsWriteOut == true)
				{
					//Write Results to File
					for (j = 0; j<n; j++) //Fit parameters
						Solver.obj.Model.MonteCarlo.MonteResults(SimCnt - 1, j) = V(j);
					_fprintf_p(MCout, "%d", SimCnt);
					_fprintf_p(MCout, "%s", Delim);
					Cov = InvertMatrix(ATA);
					for (i = 0; i<(2 * n); i += 2)
					{
						_fprintf_p(MCout, "%g", V(i / 2));
						_fprintf_p(MCout, "%s", Delim);
						_fprintf_p(MCout, "%g", sqrt(Cov(i / 2, i / 2)));
						_fprintf_p(MCout, "%s", Delim);
					}
					_fprintf_p(MCout, "%g", ChiSqr);
					_fprintf_p(MCout, "%s", Delim);

					DoF = (double)TracerNum - n;  //degrees of freedom
					if (DoF>0 && ChiSqr<100000)
					{
						boost::math::chi_squared_distribution<> ChiSqrDist(DoF);
						_fprintf_p(MCout, "%g", 1 - cdf(ChiSqrDist, ChiSqr));
					}
					else
					{
						_fprintf_p(MCout, "%g", (double)-99);
					}
					_fprintf_p(MCout, "%s", Delim);

					_fprintf_p(MCout, "%g", (double)HiTracer);
					_fprintf_p(MCout, "%s", Delim);

					_fprintf_p(MCout, "%g", HiChiSqr);
					_fprintf_p(MCout, "%s", Delim);

					_fprintf_p(MCout, "%g", (double)nIters);
					_fprintf_p(MCout, "%s", Delim);

					_fprintf_p(MCout, "%g", (std::clock() - start) / (double)CLOCKS_PER_SEC);
					_fprintf_p(MCout, "%s", Delim);
					for (int j = 0; j<TracerNum; j++) //Modeled Tracer Concentrations
					{
						Solver.obj.Model.MonteCarlo.MonteResults(SimCnt - 1, j + n) = InitTracerOutput(j);
						_fprintf_p(MCout, "%g", InitTracerOutput(j));
						_fprintf_p(MCout, "%s", Delim);
					}
					for (int j = 0; j<TracerNum; j++) //Simulated Tracer Concentrations
					{
						_fprintf_p(MCout, "%g", Solver.obj.Sample.MeasTracerConcs(j));
						_fprintf_p(MCout, "%s", Delim);
					}
					_fprintf_p(MCout, "%s\n", "");
				}
				else
				{
					for (j = 0; j<n; j++) //Fit parameters
					{
						Solver.obj.Model.MonteCarlo.MonteResults(SimCnt - 1, j) = V(j);
#ifdef _DEBUG
						_fprintf_p(stream, "%g", V(j));
						_fprintf_p(stream, "%s", Delim);
#endif
					}
					for (int j = 0; j<TracerNum; j++) //Modeled Tracer Concentrations
						Solver.obj.Model.MonteCarlo.MonteResults(SimCnt - 1, j + n) = InitTracerOutput(j);
#ifdef _DEBUG
					_fprintf_p(stream, "%s\n", "");
#endif
				}

				//Start Monte Carlo Simulations
				if (Solver.obj.Model.MonteCarlo.IsMonteCarlo == true && SimCnt<Solver.obj.Model.MonteCarlo.TotalSims)
				{
					ChiSqr = 0;
					PrevChiSqr = 0;
					for (int j = 0; j<TracerNum; j++)
					{
						Solver.obj.Sample.MeasTracerConcs(j) = Solver.obj.Model.MonteCarlo.SimulatedConcs(SimCnt, j);
						D(j) = (Solver.obj.Sample.MeasTracerConcs(j) - InitTracerOutput(j)) / Solver.obj.Sample.MeasSigmas(j);
						ChiSqr += D(j)*D(j);
#ifdef _DEBUG
						_fprintf_p(stream, "%s", Delim);
						_fprintf_p(stream, "%g\n", Solver.obj.Model.MonteCarlo.SimulatedConcs(SimCnt, j));
#endif
					}
					for (j = 0; j<8; j++)
						ModelArgs[j] = test = Solver.obj.Model.InitModVals(j);
#ifdef _DEBUG
					_fprintf_p(stream, "%s", Delim);
					_fprintf_p(stream, "%g\n", ChiSqr);
#endif
					//SendMessage(hwndPB, PBM_STEPIT, 0, 0);
					SendMessage(GetDlgItem(hwndPB, IDC_PROGBAR), PBM_STEPIT, 0, 0);
					UpdateWindow(hwndPB);
					SimCnt++;
				}
				else if (Solver.obj.Model.MonteCarlo.IsMonteCarlo)
					SimCnt++;
			} while (Solver.obj.Model.MonteCarlo.IsMonteCarlo == true && SimCnt <= Solver.obj.Model.MonteCarlo.TotalSims);
			if (Solver.obj.Model.MonteCarlo.IsMonteCarlo == true)
			{
				xOpArray[2 * n + 6].xltype = xltypeNum;
				xOpArray[2 * n + 6].val.num = (std::clock() - start) / (double)CLOCKS_PER_SEC;
				VectorXd StatVec = VectorXd::Zero(Solver.obj.Model.MonteCarlo.TotalSims);
				j = 2 * (n + TracerNum);
				SimCnt = 0;
				while (j>0)
				{
					StatVec = Solver.obj.Model.MonteCarlo.MonteResults.block(0, SimCnt, Solver.obj.Model.MonteCarlo.TotalSims, 1);
					Mean = StatVec.mean();
					for (int i = 0; i<Solver.obj.Model.MonteCarlo.TotalSims; i++)
						StatVec(i) = pow(Solver.obj.Model.MonteCarlo.MonteResults(i, SimCnt) - Mean, 2);
					StdDev = sqrt(StatVec.sum() / Solver.obj.Model.MonteCarlo.TotalSims);
					xOpArray[4 * n + 7 + (2 * TracerNum) - j].xltype = xltypeNum;
					xOpArray[4 * n + 7 + (2 * TracerNum) - j].val.num = Mean;
					xOpArray[4 * n + 7 + (2 * TracerNum) - j + 1].xltype = xltypeNum;
					xOpArray[4 * n + 7 + (2 * TracerNum) - j + 1].val.num = StdDev;
					j += -2;
					SimCnt++;
				}
				UnhookExcelWindow(g_hWndMain);
				Excel12f(xlDisableXLMsgs, 0, 0);
				DestroyWindow(hwndPB);
				if (Solver.obj.Model.MonteCarlo.IsWriteOut == true)
				{
					fclose(MCout);
				}
			}
			// Create an array of pointers to XLOPER12 values.
			LPXLOPER12 xArray = (LPXLOPER12)malloc(size * sizeof(LPXLOPER12));
			xArray->xltype = xltypeMulti | xlbitDLLFree;
			xArray->val.array.columns = 1;
			xArray->val.array.rows = size;
			xArray->val.array.lparray = xOpArray;

#ifdef _DEBUG
			LPXLOPER12 px;
			_fprintf_p(stream, "%s\n", "");
			for (i = 0; i<2 * n + 6; i++)
			{
				px = xArray->val.array.lparray + i;
				_fprintf_p(stream, "%g\n", px->val.num);
				//_fprintf_p( stream, "%s", Delim );
			}
			fclose(stream);
#endif
			return xArray;
		}
		else
		{
			xMulti[0].val.num = 0;
			xMulti[0].xltype = xltypeNum;
			return (LPXLOPER12)&xMulti[0];
		}
}


__declspec(dllexport) LPXLOPER12 WINAPI UZ_Solver1D(double &delta_time, double &TotSimTime, double &delta_depth, double &max_depth,
	const double &effective_diffusion, const double &effective_velocity,
	const double &decay_rate, const double &requested_depth, LPXLOPER12 DateRange, LPXLOPER12 surface_tracer_concs)
{
	LPXLOPER12 px, py;
	
	px = DateRange->val.array.lparray;
	py = DateRange->val.array.lparray + DateRange->val.array.rows - 1;
	bool IsDescending;
	if (py->val.num < px->val.num)
		IsDescending = true;
	else
		IsDescending = false;

	if (delta_time<=0 || TotSimTime <= 0 || delta_depth <= 0 || max_depth <= 0 || effective_diffusion <= 0 || effective_velocity <= 0 || TotSimTime <= delta_time || max_depth <= delta_depth || 
		requested_depth < 0 || surface_tracer_concs->xltype != xltypeMulti)
		return 0;

	// Create an array of XLOPER12 values.
	/*XLOPER12 *xOpArray = (XLOPER12 *)malloc((time_steps + 1) * sizeof(XLOPER12));

	for (int i = 0; i < time_steps + 1; i++)
	{
		xOpArray[i].xltype = xltypeNum;
		xOpArray[i].val.num = -99;
	}*/

	return FullyImplicitAtDepthXL(delta_time, TotSimTime, delta_depth, max_depth, effective_diffusion, effective_velocity,
		decay_rate, requested_depth, surface_tracer_concs, IsDescending);
}

__declspec(dllexport) LPXLOPER12 WINAPI UZ_Solver2D(double &delta_time, double &TotSimTime, double &delta_depth, double &max_depth,
	const double &effective_diffusion, const double &effective_velocity,
	const double &decay_rate, LPXLOPER12 DateRange, LPXLOPER12 surface_tracer_concs)
{
	LPXLOPER12 px, py;
	int NumTimeSteps = (int)TotSimTime/delta_time;
	int NumDepthSteps = (int)max_depth / delta_depth + 1;
	
	px = DateRange->val.array.lparray;
	py = DateRange->val.array.lparray + DateRange->val.array.rows - 1;
	bool IsDescending;
	if (py->val.num < px->val.num)
		IsDescending = true;
	else
		IsDescending = false;

	if (delta_time <= 0 || TotSimTime <= 0 || delta_depth <= 0 || max_depth <= 0 || effective_diffusion <= 0 || effective_velocity <= 0 || TotSimTime <= delta_time || max_depth <= delta_depth ||
		surface_tracer_concs->xltype != xltypeMulti)
		return 0;

	return FullyImplicitXL(delta_time, TotSimTime, delta_depth, max_depth, effective_diffusion, effective_velocity,
		decay_rate, surface_tracer_concs, IsDescending);
/*#ifdef _DEBUG
	if (fopen_s(&stream, "UZoutput.txt", "w") == 0)
	{
		int incrt;
		for (int i = 0; i < NumTimeSteps; i++)
		{
			for (int j = 0; j < NumDepthSteps; j++)
			{
				incrt = i*(NumDepthSteps) + j;
				px = solution->val.array.lparray + incrt;
				_fprintf_p(stream, "%g", px->val.num);
				_fprintf_p(stream, "%s", Delim);
			}
			_fprintf_p(stream, "%s\n", "");
		}
	}
	fclose(stream);
#endif*/
}