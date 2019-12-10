#pragma once
#ifndef LPM_H
#define LPM_H
#include <windows.h>
#include <CommCtrl.h>
#include <xlcall.h>
#include <framewrk.h>
#include <time.h>
#include <Eigen\Core>
#include <Eigen\Dense>
#include <math.h>
#include <omp.h>
//#include "complexMod"
//#include "cmathMod"
#include <iostream>
#include <string>
#include <stdlib.h>

#include <boost/math/distributions/gamma.hpp>
#include <boost/math/distributions/chi_squared.hpp>
#include <boost/random.hpp>
#include <boost/random/normal_distribution.hpp>
//#include "unsaturated_zone_gas_tracer-master\unsaturated_zone_tracer_solver.h"

using namespace Eigen;
using namespace std; 
using namespace boost::math;

//namespace THM = unsaturated_zone_tracer_solver_internal;
#pragma inline_recursion ( on )
#pragma intrinsic( cos )
#pragma intrinsic( fabs )
#pragma intrinsic( pow )
#pragma intrinsic( exp )

//
//------------------------------------LPM CLASS----------------------------------------------
//


static const double Tol= 1E-06;        // Stopping criteria for LPM output 
static const double Udecay = 1.19E-13;
static const double THdecay = 2.88E-14;
static const double PI = 3.1415926535897932384626433832795028841971693993751;
static const double RndNum[5] = {0.0521,0.2311,0.00125,0.4860,0.8913};
static const double glX[3] = {0.94288241569547971905635175843185720232,0.64185334234578130578123554132903188354,0.23638319966214988028222377349205292599};
static const double glA[7] = {0.015827191973480183087169986733305510591,0.094273840218850045531282505077108171960,0.15507198733658539625363597980210298680,
				0.18882157396018245442000533937297167125, 0.19977340522685852679206802206648840246, 0.22492646533333952701601768799639508076, 0.24261107190140773379964095790325635233};
static const double Roots11[11] = {0.978228658146057,0.887062599768095,0.730152005574049,0.519096129206811,0.269543155952345,0,-0.269543155952345,-0.519096129206811,-0.730152005574049,-0.887062599768095,-0.978228658146057};
static const double Coeff11[11] = {0.0556685671161737,0.125580369464904,0.186290210927734,0.23319376459199,0.262804544510246,0.2729250867779,0.262804544510246,0.23319376459199,0.186290210927734,0.125580369464904,0.0556685671161737};
static const double Roots7[7] = {0.9491079123427585,0.7415311855993945,0.4058451513773972,0.0,-0.4058451513773972,-0.7415311855993945,-0.9491079123427585};
static const double Coeff7[7] = {0.1294849661688697,0.2797053914892766,0.3818300505051189,0.4179591836734694,0.3818300505051189,0.2797053914892766,0.1294849661688697};
static const double Roots5[5] = {0.906179845938664,0.538469310105683,0.0,-0.538469310105683,-0.906179845938664};
static const double Coeff5[5] = {0.236926885056189,0.478628670499366,0.568888888888889,0.478628670499366,0.236926885056189};
static const double Roots3[3] = {sqrt(0.6),0.0,-sqrt(0.6)};
static const double Coeff3[3] = {double(5)/double(9),double(8)/double(9),double(5)/double(9)};
static const double Sqrt1=0.81649658092772603, Sqrt2=0.44721359549995793;

#ifdef _DEBUG
	FILE    *stream = NULL;
	char    Delim[3] = ", ";
#endif
MatrixXd InvertMatrix(MatrixXd MatrixToInvert);
VectorXd SolveSVD(MatrixXd Xmatrix, VectorXd Bvec);
void cwCenter(HWND, int);
INT_PTR CALLBACK DIALOGMsgProc(HWND hWndDlg, UINT message, WPARAM wParam, LPARAM lParam);
BOOL GetHwnd(HWND * pHwnd);
int lpwstricmp(LPWSTR s, LPWSTR t);
extern void FAR PASCAL HookExcelWindow(HANDLE hWndExcel);
extern void FAR PASCAL UnhookExcelWindow(HANDLE hWndExcel);
LRESULT CALLBACK WndProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
HWND CreateProgressBase();
void WINAPI xlAutoFree12(LPXLOPER12 pxFree);

class LPM
{
private:	


public:
	struct LPM_Object
	{
		struct Sample
		{
			VectorXd SampleDates;
			VectorXd MeasTracerConcs;
			VectorXd MeasSigmas;
			MatrixXd ActiveVals;
		}Sample;
		struct Model
		{
			int ModelNum;
			struct FitParmIndexes
			{
				vector<int> Val;
				bool isUZtime;
			}FitParmIndexes;

			VectorXd InitModVals;
			VectorXd HiBounds;
			VectorXd LowBounds;
			struct MonteCarlo
			{
				bool IsMonteCarlo;
				VectorXd OrigMeasVals;
				MatrixXd MonteResults;
				MatrixXd SimulatedConcs;
				int TotalSims;
				bool IsWriteOut;
			}MonteCarlo;
		}Model;
		struct Tracer
		{
			vector<int> Tracers;
			VectorXd DateRange;
			MatrixXd TracerInputRange;
			VectorXd Lambda;
			VectorXd UZtime;
			vector<int> UZtimeCond;
			vector<int> CalcCond;
			struct TracerComp_2
			{
				MatrixXd Val;
				bool isValid;
			}TracerComp_2;
			struct He4
			{
				double Uppm;
				double THppm;
				double Porosity;
				double SedDensity;
				double He4SolnRate;
			}He4;
		}Tracer;
		//name?
		int k;
		int k1;
		int k2;
		double Lambda2[10];	
		
	}obj;
		
	LPM ();
	LPM (int, LPXLOPER12, LPXLOPER12 , LPXLOPER12 , LPXLOPER12, FP*, LPXLOPER12, LPXLOPER12 , LPXLOPER12, LPXLOPER12, LPXLOPER12,
		LPXLOPER12, LPXLOPER12, LPXLOPER12, LPXLOPER12, double, double, double, double, double, LPXLOPER12, int, LPXLOPER12);


	int LPM::sign(const double X)  const // returns 0 for 0, 1 for positive, -1 for negative.  Faster than abs()
	{   
		return ((X==0)?0:(X>0)?1:-1); 
	}

	double LPM::AlphaDensityInt(double T, double& X, double& Alpha) const
	{
		return cos(X*T) * exp(-pow((T>=0)?T:-1*T,Alpha));
	}

	double LPM::ReturnEndDate(double SampleDate, double TimeIncrement)
	{
		return floor (SampleDate) + floor ((SampleDate-floor (SampleDate)) / TimeIncrement) * TimeIncrement; 
	}

	double LPM::ReturnLambdaCorrection(double MinAge, double MaxAge, double Lambda) //returns Lambda Correction Factor
	{
		return exp(-Lambda*(MinAge+(MaxAge-MinAge)/2));  
		//Above solution works well for small time steps <1 yr
		//Below is the exact integral equation for the interval
		//return (exp(-Lambda*MinAge)-exp(-Lambda*MaxAge))/Lambda/(MaxAge-MinAge);
	}

	double LPM::GetRandomDoubleUsingNormalDistribution(double mean, double sigma)
	{
		typedef boost::normal_distribution<double> NormalDistribution;
		typedef boost::mt19937 RandomGenerator;
		typedef boost::variate_generator<RandomGenerator&, \
								NormalDistribution> GaussianGenerator;
 
		/** Initiate Random Number generator with current time */
		static RandomGenerator rng(static_cast<unsigned> (time(0)));
 
		/* Choose Normal Distribution */
		NormalDistribution gaussian_dist(mean, sigma);
 
		/* Create a Gaussian Random Number generator
		*  by binding with previously defined
		*  normal distribution object
		*/
		GaussianGenerator generator(rng, gaussian_dist);
 
		// sample from the distribution
		return generator();
	}

	VectorXd LPM_TracerOutput(double, double, double, double, double, double, double, double, double);
	VectorXd d_dx_LPM_Model(double, double, double, double, double, double, double, double, double, int);

	VectorXd PFM_Int(double, double);
	VectorXd DM_Internal(double, double, double);
	VectorXd PEM_Int(double, double, double, double);
	VectorXd GAM_Int(double, double, double);
	VectorXd EMM_Int(double, double);
	VectorXd EPM_Int(double, double,double);
	VectorXd FDM_Internal(double, double, double, double);

	double gt_DMint(double, double, double, double, double);
	double gt_FDMint(double, double&, double&, double&, double&, double);	
	double gt_PFM(double, double, double, double);

	double DispInt(double, double, double, double); //Calculates dispersion []?, called by gt_DMint
	double AlphaStablePDFauxGL(double, double, double, double, double, double&, double&) ;
	double AlphaStablePDFInt(double, double&, double&, double&, double);
	
	void TracerParseEM(double, double,  int&, double, double[], int[], int, double&, double&, int, int);
	void TracerParseDM(double, double, int&, double, double[], int[], int);
	std::vector<double> ThomasAlgorithimSingleValue(
		const double lower_diagonal, const double middle_diagonal,
		const double upper_diagonal,
		const std::vector<double> rhs_vector);
	LPXLOPER12 FullyImplicitAtDepthXL(
		const double delta_time, const double TotSimTime, const double delta_depth, const double max_depth,
		const double effective_diffusion, const double effective_velocity, const double decay_rate, const double requested_depth,
		LPXLOPER12 &surface_tracer_concentrations, bool IsDescending);
	LPXLOPER12 FullyImplicitXL(
		const double delta_time, const double TotSimTime, const double delta_depth, const double max_depth,
		const double effective_diffusion, const double effective_velocity, const double decay_rate,
		LPXLOPER12 &surface_tracer_concentrations, bool IsDescending);
};


LPM::LPM()
{
	
	//obj.Model.FitParmIndexes.Val.resize(0);
	//obj.Sample.MeasTracerConcs.resize(1);
	obj.Model.InitModVals.resize(1);
	obj.Tracer.DateRange.resize(1);
	obj.Tracer.Lambda.resize(1);
	obj.Sample.SampleDates.resize(1);
	obj.Tracer.UZtime.resize(1);
	//obj.Sample.MeasSigmas.resize(1);
	obj.Model.HiBounds.resize(1);
	obj.Model.LowBounds.resize(1);
	obj.Tracer.TracerComp_2.Val.resize(1,1);
	
}


LPM::LPM(int ModelNum, LPXLOPER12 lxMeasTracerConcs, LPXLOPER12 lxFitParmIndexes, LPXLOPER12 lxInitModVals, LPXLOPER12 lxTracers, FP lxdateRange[], 
	LPXLOPER12 lxTracerInputRange, LPXLOPER12 lxLambda, LPXLOPER12 lxSampleDates, 
	LPXLOPER12 lxuzTime, LPXLOPER12 lxUZtimeCond, LPXLOPER12 lxMeasSigmas, LPXLOPER12 lxHiBounds, LPXLOPER12 lxLowBounds,
	LPXLOPER12 lxTracerComp_2, double Uppm, double THppm, double Porosity, double SedDensity, double He4SolnRate, LPXLOPER12 lxIsMonteCarlo, int iTotalSims, LPXLOPER12 lxIsWriteOut)
{
	int size=0, i=0, n, k, j, rows, cols, isHe4=0, TracerBool=0;
	LPXLOPER12 px, py;			// Pointer into array 

	try
	{
	
		obj.Model.InitModVals.resize(1);
		obj.Tracer.DateRange.resize(1);
		obj.Tracer.Lambda.resize(1);
		obj.Tracer.UZtime.resize(1);
		obj.Model.HiBounds.resize(1);
		obj.Model.LowBounds.resize(1);
		obj.Tracer.TracerComp_2.Val.resize(1,1);
		obj.Model.ModelNum=ModelNum;

		//SampleDates
		if(lxSampleDates->xltype!=xltypeMulti)
		{
			obj.Sample.SampleDates.resize(1);
			obj.Sample.SampleDates(0)=lxSampleDates->val.num;
		}
		else
		{
			size=lxSampleDates->val.array.columns*lxSampleDates->val.array.rows;
			obj.Sample.SampleDates.resize(size);
			for(i=0;i<size;i++)
			{
				px = lxSampleDates->val.array.lparray + i;
				obj.Sample.SampleDates(i)=px->val.num;
			}
		}
	
		//MeasTracerConcs
	
		if((lxMeasTracerConcs->xltype && lxMeasSigmas->xltype) == xltypeNum && lxMeasTracerConcs->val.array.rows == lxMeasSigmas->val.array.rows
			&& lxMeasTracerConcs->val.array.columns == lxMeasSigmas->val.array.columns)
		{
	
			if(lxMeasTracerConcs->xltype!=xltypeMulti)
			{
				obj.Sample.MeasTracerConcs.resize(1,1);
				obj.Sample.MeasTracerConcs(0,0)=lxMeasTracerConcs->val.num;
				obj.Sample.ActiveVals.resize(1,1);
				obj.Sample.ActiveVals(0,0)=1;

				obj.Sample.MeasSigmas.resize(1,1);
				obj.Sample.MeasSigmas(0,0)=lxMeasSigmas->val.num;
			}
			else
			{
				vector<double> MeasTracerHold;
				vector<double> MeasSigmaHold;
				
				obj.Sample.ActiveVals = MatrixXd::Zero(lxMeasTracerConcs->val.array.rows,lxMeasTracerConcs->val.array.columns);
				k=0;
				for(i=0;i<lxMeasTracerConcs->val.array.rows;i++)
				{
					for(j=0;j<lxMeasTracerConcs->val.array.columns;j++)
					{
						px = lxMeasTracerConcs->val.array.lparray + k;
						py = lxMeasSigmas->val.array.lparray + k;

						if(px->xltype==xltypeNum && px->val.num>=0 && py->xltype==xltypeNum && py->val.num>=0)
						{
							MeasTracerHold.push_back(px->val.num);
							MeasSigmaHold.push_back(py->val.num);
							obj.Sample.ActiveVals(i,j)=1;
						}
						k++;
					}
				}
				obj.Sample.MeasTracerConcs.resize(MeasTracerHold.size());
				obj.Sample.MeasSigmas.resize(MeasSigmaHold.size());
				for(i=0;i<MeasTracerHold.size();i++)
				{
					obj.Sample.MeasTracerConcs(i)=MeasTracerHold[i];
					obj.Sample.MeasSigmas(i)=MeasSigmaHold[i];
				}
			}
		}
		else
			throw 2;
		
	
		//Tracers
		if(lxTracers->xltype!=xltypeMulti)						//3H = 1
		{														//3He(trit) = 2
			obj.Tracer.Tracers.resize(1);						//3Ho = 3
			n=lxTracers->val.num;								//3H/3Ho = 4
			if(n==5)											//He4 = 5
				isHe4=1;										//14C = 5
			obj.Tracer.Tracers[0]=n;							
		}
		else
		{
			size=lxTracers->val.array.columns*lxTracers->val.array.rows;
			obj.Tracer.Tracers.resize(size);
			for(i=0;i<size;i++)
			{
				n=obj.Tracer.Tracers[i]=lxTracers->val.array.lparray[i].val.num;
				if(n==5)
					isHe4=1;
			}
		}
	
		//He4
		if(isHe4)
		{
			obj.Tracer.He4.Uppm=Uppm;
			obj.Tracer.He4.THppm=THppm;
			obj.Tracer.He4.Porosity=Porosity;
			obj.Tracer.He4.SedDensity=SedDensity;
			obj.Tracer.He4.He4SolnRate=He4SolnRate;
		}

		//Lambda
		if(lxLambda->xltype!=xltypeMulti)
			obj.Tracer.Lambda(0)=lxLambda->val.num;
		else
		{
			size=lxLambda->val.array.columns*lxLambda->val.array.rows;
			obj.Tracer.Lambda.resize(size);
			for(i=0;i<size;i++)
				obj.Tracer.Lambda(i)=lxLambda->val.array.lparray[i].val.num;
		}

		//uzTime
		if(lxuzTime->xltype!=xltypeMulti)
			obj.Tracer.UZtime(0)=lxuzTime->val.num;
		else
		{
			size=lxuzTime->val.array.columns*lxuzTime->val.array.rows;
			obj.Tracer.UZtime.resize(size);
			for(i=0;i<size;i++)
				obj.Tracer.UZtime(i)=lxuzTime->val.array.lparray[i].val.num;
		}

		//UZtimeCond
		if(lxUZtimeCond->xltype!=xltypeMulti)
			obj.Tracer.UZtimeCond.push_back(lxUZtimeCond->val.num);
		else
		{
			size=lxUZtimeCond->val.array.columns*lxUZtimeCond->val.array.rows;
			for(i=0;i<size;i++)
				obj.Tracer.UZtimeCond.push_back(lxUZtimeCond->val.array.lparray[i].val.num);
		}

		//CalculateCond
		for(k=0;k<obj.Tracer.Tracers.size();k++)
		{
			switch(obj.Tracer.Tracers[k])
			{
			case 1://3H
				obj.Tracer.CalcCond.push_back(1);
				obj.k1=k;
				obj.Lambda2[k]=obj.Tracer.Lambda(k);
				break;
			case 2://3He(trit)
				obj.Tracer.CalcCond.push_back(2);
				if(obj.Tracer.Lambda(k) ==0)
					obj.Lambda2[k]=log(2.0)/12.32;
				break;
			case 3://3Ho
				obj.Tracer.CalcCond.push_back(1);
				obj.k2=k;
				if(obj.Tracer.Lambda(k) ==0)
					obj.Lambda2[k]=log(2.0)/12.32;
				break;
			case 4://3H/3Ho
				obj.Tracer.CalcCond.push_back(3);
				if(obj.Tracer.Lambda(k) ==0)
					obj.Lambda2[k]=log(2.0)/12.32;
				break;
			case 5: //He4
				obj.Tracer.CalcCond.push_back(4);
				obj.Lambda2[k]=obj.Tracer.Lambda(k);
				break;
			case 6://14C
				obj.Tracer.CalcCond.push_back(1);
				obj.Lambda2[k]=obj.Tracer.Lambda(k);
				break;
			default:
				obj.Tracer.CalcCond.push_back(1);
				obj.Lambda2[k]=obj.Tracer.Lambda(k);
				break;
			}
		}

		//FitParmIndexes
		obj.Model.FitParmIndexes.isUZtime=false;
		obj.Model.FitParmIndexes.Val.resize(0);
		for(i=0;i<lxFitParmIndexes->val.array.columns*lxFitParmIndexes->val.array.rows;i++)
		{
			n=lxFitParmIndexes->val.array.lparray[i].val.num;
			if(n>0)
				obj.Model.FitParmIndexes.Val.push_back(n);		
			if(n==1)
				obj.Model.FitParmIndexes.isUZtime=true;
		}
		
		if (lxIsMonteCarlo->val.xbool ==true)
		{
			obj.Model.MonteCarlo.IsMonteCarlo =true;
			obj.Model.MonteCarlo.TotalSims = iTotalSims;
			obj.Model.MonteCarlo.OrigMeasVals = obj.Sample.MeasTracerConcs;
			obj.Model.MonteCarlo.SimulatedConcs = MatrixXd::Zero(obj.Model.MonteCarlo.TotalSims,obj.Sample.MeasTracerConcs.size());
			n=obj.Model.FitParmIndexes.Val.size();
			obj.Model.MonteCarlo.MonteResults = MatrixXd::Zero(obj.Model.MonteCarlo.TotalSims, n + obj.Sample.MeasTracerConcs.size()); //Remember that n is set in the last for statement
			double junk = 0.0;
			//junk = obj.Model.MonteCarlo.TotalSims;
			//junk = obj.Model.MonteCarlo.OrigMeasVals(0);
			for (int i=0;i<obj.Model.MonteCarlo.TotalSims; i++)
			{
				for (int j=0; j<obj.Sample.MeasTracerConcs.size(); j++)
				{
					junk = GetRandomDoubleUsingNormalDistribution(obj.Model.MonteCarlo.OrigMeasVals(j), obj.Sample.MeasSigmas(j));
					if (junk<0)
						obj.Model.MonteCarlo.SimulatedConcs(i,j) = 0.0;
					else
						obj.Model.MonteCarlo.SimulatedConcs(i,j) = junk;
				}
			}
			if (lxIsWriteOut->val.xbool ==true)
				obj.Model.MonteCarlo.IsWriteOut =true;
		}
		else
		{
			obj.Model.MonteCarlo.IsMonteCarlo =false;
			obj.Model.MonteCarlo.IsWriteOut =false;
		}

		//InitModVals
		if(lxInitModVals->xltype!=xltypeMulti)
			obj.Model.InitModVals(0)=lxInitModVals->val.num;
		else
		{
			size=lxInitModVals->val.array.columns*lxInitModVals->val.array.rows;
			obj.Model.InitModVals.resize(size);
			for(i=0;i<size;i++)
				obj.Model.InitModVals(i)=lxInitModVals->val.array.lparray[i].val.num;
		}

		size=lxdateRange->rows;
		obj.Tracer.DateRange.resize(size);    //date range
		for(i=0;i<size;i++)
			obj.Tracer.DateRange(i)=lxdateRange->array[i];

		//double X;
		rows=lxTracerInputRange->val.array.rows;		//tracer input range
		cols=lxTracerInputRange->val.array.columns;
		//px=lxTracerInputRange->val.array.lparray;
		//Map<Matrix<int,2,4,RowMajor> >
		//MatrixXd eigenX(rows,cols);
		//eigenX= Map<MatrixXd, Unaligned,Stride<1,4> >( &px->val.num, rows, cols );
		obj.Tracer.TracerInputRange.resize(rows, cols);
		//obj.Tracer.TracerInputRange = Map<MatrixXd>( &px->val.num, lxTracerInputRange->val.array.rows, lxTracerInputRange->val.array.columns );
		
		// obtain a pointer to the current item //
		k=0;
		for(i=0;i<rows;i++)
		{
			for(n=0;n<cols;n++)
			{
				px = lxTracerInputRange->val.array.lparray + k;
				obj.Tracer.TracerInputRange(i,n)=px->val.num;
				//X=eigenX(i,n);
				//if(X>0)
				//	X=X;
				k++;
			}
		}

		 //HiBounds
		if(lxHiBounds->xltype!=xltypeMulti)
			obj.Model.HiBounds(0)=lxHiBounds->val.num;
		else
		{
			size=lxHiBounds->val.array.columns*lxHiBounds->val.array.rows;
			obj.Model.HiBounds.resize(size);
			for(i=0;i<size;i++)
			{
				px = lxHiBounds->val.array.lparray + i;
				obj.Model.HiBounds(i)=px->val.num;
			}
		}

	
		 //LowBounds
		if(lxLowBounds->xltype!=xltypeMulti)
			obj.Model.LowBounds(0)=lxLowBounds->val.num;
		else
		{
			size=lxLowBounds->val.array.columns*lxLowBounds->val.array.rows;
			obj.Model.LowBounds.resize(size);
			for(i=0;i<size;i++)
			{
				px = lxLowBounds->val.array.lparray + i;
				obj.Model.LowBounds(i)=px->val.num;
			}
		}

		 //TracerComp_2
		obj.Tracer.TracerComp_2.isValid=true;
		if(lxTracerComp_2->xltype!=xltypeMulti)
		{
			obj.Tracer.TracerComp_2.Val(0,0)=lxTracerComp_2->val.num;
			if(obj.Tracer.TracerComp_2.Val(0,0)==0)
				obj.Tracer.TracerComp_2.isValid=false;
		}
		else
		{
			obj.Tracer.TracerComp_2.Val.resize(lxTracerComp_2->val.array.rows,lxTracerComp_2->val.array.columns);
			for(i=0;i<lxTracerComp_2->val.array.rows;i++)
			{
				for(int j=0;j<lxTracerComp_2->val.array.columns;j++)
					{
						obj.Tracer.TracerComp_2.Val(i,j)=lxTracerComp_2->val.array.lparray[i+j].val.num;
						if(obj.Tracer.TracerComp_2.Val(i,j)>0)
							TracerBool++;
				}
			}
		if(TracerBool==0)
			obj.Tracer.TracerComp_2.isValid=false;
		}
	}
	catch (int error)
	{
		XCHAR szBuf[255];
		wsprintfW((LPWSTR)szBuf, L"Exception: %i", error, L"\n" ,
			L"Non-matching measured tracer concs and sigmas");
		Excel12f(xlcAlert, 0, 2, TempStr12(szBuf), TempInt12(2));
	}
}


VectorXd LPM::PFM_Int(double Tau, double SampleDate)
{
	VectorXd DMprev, Result;
	double DR, Cin,MaxAge=0.0,PFMprev[10];
	int j, StepInc, StopCriteria,NameCheck=0,nIters = 2000000;
	int TracerCount=0, CheckComplete=0, TracerNum=0, Js[10];

	TracerNum=obj.Tracer.Tracers.size();
	Result.resize(TracerNum);
	
	j = obj.Tracer.DateRange.rows() - 1;
	StepInc = 1;
	StopCriteria = 0;
	if (obj.Tracer.DateRange(j) < obj.Tracer.DateRange(0))
	{
		StepInc = -1;
		StopCriteria = j;
		j = 0;
	}
	DR = obj.Tracer.DateRange(j);
	
	for (int k = 0;k < TracerNum; k++)
	{
		Js[k] = j;
		DR = obj.Tracer.DateRange(Js[k]);
		while ((SampleDate - Tau - obj.Tracer.UZtime(k) < DR) && Js[k] != StopCriteria)
		{
			Js[k] = Js[k] - StepInc;
			DR = obj.Tracer.DateRange(Js[k]);
		}			
	}

	for(int k=0;k<TracerNum;k++)
	{
		Cin = obj.Tracer.TracerInputRange(Js[k],k);
		if (Cin == 0 && Js[k] == StopCriteria && obj.Tracer.CalcCond[k] != 4)
		{
			return Result;
		}

		switch(obj.Tracer.CalcCond[k])
		{
		case 1:
			Result(k) = Cin * exp(-obj.Lambda2[k]* obj.Tracer.UZtime(k))*exp(-obj.Tracer.Lambda(k) * Tau);
			break;
		case 2:
			Cin*=(1-exp(-obj.Lambda2[k]*Tau));
			Result(k)= Cin * exp(-obj.Lambda2[k]* obj.Tracer.UZtime(k));
			break;
		case 3:
			if(obj.k1==0 && obj.k2==0)
				Result(k)=exp(-obj.Lambda2[k]* Tau)/exp(-obj.Tracer.Lambda(k) * Tau);
			else if(obj.k1!=0 || obj.k2!=0)
				Result(k)=Result(obj.k1)/Result(obj.k2);
			break;
		case 4:
			if (obj.Tracer.He4.He4SolnRate != 0.0)
				Result(k) = obj.Tracer.He4.He4SolnRate * Tau;
			else
				Result(k) = obj.Tracer.He4.SedDensity / obj.Tracer.He4.Porosity * (Udecay * obj.Tracer.He4.Uppm + THdecay * obj.Tracer.He4.THppm) * Tau;
			break;
		default:
			break;
		}
		PFMprev[k]=Result(k);
	}
	return Result;
}

VectorXd LPM::DM_Internal(double Tau, double SampleDate, double DP)
{
	double DMprev[10], EndDates[10];
	double DR, Cin; 
	double EndDate=0,  DMnoDecay=0.0, TimeIncrement, MaxAge=0.0, Integral=0.0,gtZeroLam=0.0, CummFrac=0.0, InitTritResult=0.0;
	int i,j, nIters=1000000, StepInc, StopCriteria, NameCheck=0, Js[10];
	int TracerCount=0, CheckComplete=0, k, TracerNum=0;	
	
	TracerNum=obj.Tracer.Tracers.size();
	VectorXd Result = VectorXd::Zero(TracerNum);
	
	if(obj.Tracer.DateRange.rows() == 1 || ( obj.Tracer.TracerInputRange.rows() == 1) || Tau ==0 || SampleDate==0 || DP == 0) 
		return Result;
    if (DP < 1.0)
        TimeIncrement = 1.0 / 12.0;
    else
        TimeIncrement = 1.0 / 12.0 / 2.0;
    if (Tau >= 100.) // reset time increment and end date for large tau
	{
        TimeIncrement = fabs(obj.Tracer.DateRange(1)-obj.Tracer.DateRange(0)); //MinTimeInc(DateRange);
		if (fabs(obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 1)-obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 2)) < TimeIncrement)
			TimeIncrement = fabs(obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 1)-obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 2));
	}

    j = obj.Tracer.DateRange.rows()-1;
	StepInc = 1;
	StopCriteria = 0;
	if (obj.Tracer.DateRange(j) < obj.Tracer.DateRange(0))
	{
		StepInc = -1;
		StopCriteria = j;
		j = 0;
	}
	DR = obj.Tracer.DateRange(j);
	TracerParseDM(EndDate,TimeIncrement,TracerCount,SampleDate,EndDates,Js,j);

	for (i = 1; i<= nIters; i++)
	{
		MaxAge = i * TimeIncrement;
		gtZeroLam = LPM::gt_DMint(MaxAge - TimeIncrement, MaxAge, Tau, 0.0, DP);
		CummFrac+=gtZeroLam;
		CheckComplete=0;
		//loop through tracers
		for(k=0;k<TracerNum;k++)
		{
			Integral = gtZeroLam*ReturnLambdaCorrection(MaxAge - TimeIncrement, MaxAge, obj.Tracer.Lambda[k]);				
			DR = obj.Tracer.DateRange(Js[k]);
			while ((EndDates[k] - MaxAge) < DR && Js[k] != StopCriteria)
			{
				Js[k] -= StepInc;
				DR = obj.Tracer.DateRange(Js[k]);
			}
			Cin = obj.Tracer.TracerInputRange(Js[k],k);
			if(obj.Tracer.Tracers[k]==5)
			{
				if (obj.Tracer.He4.He4SolnRate != 0.0)
					Cin = obj.Tracer.He4.He4SolnRate;
				else
					Cin = obj.Tracer.He4.SedDensity / obj.Tracer.He4.Porosity * (Udecay * obj.Tracer.He4.Uppm + THdecay * obj.Tracer.He4.THppm);
			}
			if (Cin > 0)
			{
				DMprev[k] = Result(k);
				switch(obj.Tracer.CalcCond[k])
				{
				case 1:  //3H, 3Ho, !3He(trit), !3H/3Ho
					Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Integral;
					break;
				case 2: //3He(trit)
					Cin*=(1-exp(-obj.Lambda2[k]*(MaxAge-TimeIncrement/2)));
					Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Integral;	
					break;
				case 3: //3H/3Ho
					if(obj.k1==0 && obj.k2==0)
					{
						DMnoDecay += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Integral;
						Integral = gt_DMint(MaxAge - TimeIncrement, MaxAge, Tau, obj.Lambda2[k], DP);
						InitTritResult+= Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) *Integral;
						Result(k) = InitTritResult/DMnoDecay;
					}
					else
						Result(k)=Result(obj.k1)/Result(obj.k2);
					break;
				case 4:
					Result(k) += Cin  * Integral * MaxAge;
					break;
				}
				if (MaxAge > Tau && Result(k) > 0) //This has to be outer if,then; otherwise Overflow occurs because of division by zero
				{
					if ((Result(k) - DMprev[k]) / Result(k) < Tol)
						CheckComplete++;
				}
			}
			else //if (Cin == 0 || Js[k] == StopCriteria)
				CheckComplete++;
		}
		if(CheckComplete==TracerCount)
			return Result;
	}
	return Result;
}

VectorXd LPM::FDM_Internal(double Tau, double Alpha, double SampleDate, double DP)
{
	double DMprev[10], EndDates[10];
	double DR, Cin; // pointers to Date Range and Tracer Input
	double EndDate=0, DMnoDecay=0.0, TimeIncrement, MaxAge=0.0, Integral=0.0,gtZeroLam=0.0, CummFrac=0.0, InitTritResult=0.0;
	int i,j, nIters, StepInc, StopCriteria, NameCheck=0, Js[10];
	int TracerCount=0, CheckComplete=0, k, TracerNum=0;

	TracerNum=obj.Tracer.Tracers.size();
	VectorXd Result = VectorXd::Zero(TracerNum);


	if(obj.Tracer.DateRange.rows() == 1 || ( obj.Tracer.TracerInputRange.cols() == 1 || obj.Tracer.TracerInputRange.rows() == 1) || Tau ==0 || SampleDate==0 || DP == 0 || Alpha==0) 
		return Result;
	
	nIters = 1000000;

	TimeIncrement = 1.0 / 12.0;
	if (Tau >= 100.) // reset time increment and end date for large tau
	{
        TimeIncrement = abs(obj.Tracer.DateRange(1)-obj.Tracer.DateRange(0));
		if (abs(obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 1)-obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 2)) < TimeIncrement)
			TimeIncrement = abs(obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 1)-obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 2));
	}

    j = obj.Tracer.DateRange.rows()-1;
	StepInc = 1;
	StopCriteria = 0;
	if (obj.Tracer.DateRange(j) < obj.Tracer.DateRange(0))
	{
		StepInc = -1;
		StopCriteria = j;
		j = 0;
	}
	DR = obj.Tracer.DateRange(j);
	// Loop through obj.Tracer.Tracers and count how many zero lambdas
	TracerParseDM(EndDate,TimeIncrement,TracerCount,SampleDate,EndDates,Js,j);

	for (i = 1; i<= nIters; i++)
	{
		MaxAge = i * TimeIncrement;
		gtZeroLam = gt_FDMint(MaxAge - TimeIncrement, MaxAge, Alpha, Tau, DP,0);
		CummFrac+=gtZeroLam;
		CheckComplete=0;
		//loop through obj.Tracer.Tracers
		for(k=0;k<TracerNum;k++)
		{
			Integral = gtZeroLam*ReturnLambdaCorrection(MaxAge - TimeIncrement, MaxAge, obj.Tracer.Lambda(k));				
			DR = obj.Tracer.DateRange(Js[k]);
			while ((EndDates[k] - MaxAge) < DR && Js[k] != StopCriteria)
			{
				Js[k] -= StepInc;
				DR = obj.Tracer.DateRange(Js[k]);
			}
			Cin = obj.Tracer.TracerInputRange(Js[k],k);
			if(obj.Tracer.Tracers[k]==5)
			{
				if (obj.Tracer.He4.He4SolnRate != 0.0)
					Cin = obj.Tracer.He4.He4SolnRate;
				else
					Cin = obj.Tracer.He4.SedDensity / obj.Tracer.He4.Porosity * (Udecay * obj.Tracer.He4.Uppm + THdecay * obj.Tracer.He4.THppm);
			}
			if (Cin > 0)
			{
				DMprev[k] = Result(k);
				switch(obj.Tracer.CalcCond[k])
				{
				case 1:
					Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Integral;
					break;
				case 2:
					Cin*=(1-exp(-obj.Lambda2[k]*(MaxAge-TimeIncrement/2)));
					Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Integral;	
					break;
				case 3:
					if(obj.k1==0 && obj.k2==0)
					{
						DMnoDecay += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Integral;
						Integral = gt_FDMint(MaxAge - TimeIncrement, MaxAge,Alpha, Tau, DP,obj.Lambda2[k]);
						InitTritResult+= Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) *Integral;
						Result(k) = InitTritResult/DMnoDecay;
					}
					else
						Result(k)=Result(obj.k1)/Result(obj.k2);
					break;
				case 4:
					Result(k) += Cin * Integral * MaxAge;
					break;
				}
				if (MaxAge > Tau && Result(k) > 0) //This has to be outer if,then; otherwise Overflow occurs because of division by zero
				{
					if ((Result(k) - DMprev[k]) / Result(k) < Tol)
					{
						CheckComplete++;
					}
				}
			}
			else //if (Cin == 0 || Js[k] == StopCriteria)
				CheckComplete++;
		}
		if(CheckComplete==TracerCount)
			return Result;
	}
	return Result;
}

void LPM::TracerParseDM(double EndDate, double TimeIncrement, int &TracerCount, double SampleDate,  //DM, FDM
	double EndDates[], int Js[], int j)
{
	for(int k=0;k<obj.Tracer.Tracers.size();k++)
	{
		if(obj.Tracer.Tracers[k]>=0)
		{
			TracerCount++;
			EndDate = ReturnEndDate(SampleDate- obj.Tracer.UZtime(k),TimeIncrement);
			if (EndDate == SampleDate - obj.Tracer.UZtime(k))
				EndDate = EndDate - 1.0/12.0;
			EndDates[k] = EndDate;
			Js[k] = j;
		}
	}
}
	
inline double LPM::DispInt(double X, double Tau, double Lambda, double DP)  //Calculates dispersion []?, called by gt_DMint
{
	double alpha, beta, phi, theta, Result3 = 0.0;
	phi = 4.0 * DP * PI;
	theta = pow ((X / Tau),3.0);
	alpha = 1.0 / (Tau * sqrt (phi) * sqrt (theta));
    beta = -(Lambda * Tau + 1.0 / (4.0 * DP)) * X / Tau - 1.0 / (4.0 * DP) * Tau / X + 1.0 / (2.0 * DP);
    Result3 = alpha * exp(beta);
    return Result3;
}

inline double LPM::gt_DMint(double MinAge, double MaxAge, double Tau, double Lambda, double DP) //Calculates dispersion [?], called by DM_Internal
{
	double Result=0.0;
    if ((MaxAge != 0) | (DP != 0))
	{
		double k1 = (MaxAge - MinAge)/2, k2 = (MinAge + MaxAge)/2;
		Result+=k1*Coeff5[0]*DispInt(k1*Roots5[0]+k2,Tau,Lambda,DP);
		Result+=k1*Coeff5[1]*DispInt(k1*Roots5[1]+k2,Tau,Lambda,DP);
		Result+=k1*Coeff5[2]*DispInt(k1*Roots5[2]+k2,Tau,Lambda,DP);
		Result+=k1*Coeff5[3]*DispInt(k1*Roots5[3]+k2,Tau,Lambda,DP);
		Result+=k1*Coeff5[4]*DispInt(k1*Roots5[4]+k2,Tau,Lambda,DP);
	}
	return Result;
}


//-----------------------------------



inline double LPM::AlphaStablePDFauxGL(double a, double b, double fa, double fb, double IS, double& X, double& Alpha) 
{   
  double m = (a + b)/2, h = (b - a)/2;
  double mll = m - Sqrt1*h, ml = m - Sqrt2*h, mr = m + Sqrt2*h, mrr = m + Sqrt1*h;
  double fmll = AlphaDensityInt(mll, X, Alpha), fml = AlphaDensityInt(ml, X, Alpha), fm = AlphaDensityInt(m, X, Alpha);
  double fmr = AlphaDensityInt(mr, X, Alpha), fmrr = AlphaDensityInt(mrr, X, Alpha);
  double i2 = (h/6)*(fa + fb + 5*(fml + fmr));
  double i1 = h/1470*(77*(fa + fb)+432*(fmll + fmrr)+625*(fml + fmr)+672*fm);
  double Eval = IS + (i1 - i2);
  if ((Eval == IS) || (mll <= a) || (b < mrr))
	return i1;
  else
    return AlphaStablePDFauxGL(a, mll, fa, fmll, IS, X, Alpha) +                    
		  AlphaStablePDFauxGL(mll, ml, fmll, fml, IS, X, Alpha) +
		  AlphaStablePDFauxGL(ml, m, fml, fm, IS, X, Alpha) +
		  AlphaStablePDFauxGL(m, mr, fm, fmr, IS, X, Alpha) +
		  AlphaStablePDFauxGL(mr, mrr, fmr, fmrr, IS, X, Alpha) +
		  AlphaStablePDFauxGL(mrr, b, fmrr, fb, IS, X, Alpha);
  
} 

inline double LPM::AlphaStablePDFInt(double T, double& Alpha, double& Tau, double& DP, double Lambda)
 {
	 //Computes the integral of an aplha stable density (with Beta = 0, Mu=0) using
	 //adaptive gauss-lobatto numerical integration (Gander and Gautschi, 2000; Adaptive Quadrature -- Revisited)
	 //with modification

	 //Modified so that the calculation is one sided (positive X) with a = 0.0 and b is dependent on alpha
	
	double Result=0.0,a=0.0, b=561*pow(Alpha, -4.583), IS=0.0, epsilon=2.22e-016,Tol = 1e-08;
	double Exponent=1/Alpha, Velocity=1/Tau, Dispersion=DP*Velocity, Var=abs(cos(PI*0.5*Alpha));
	double F1=pow(T,Exponent), F2=abs(1-T/Tau), F3=pow((DP*Var/Tau),Exponent), X=F2/(F1*F3);
	double Sigma=PI*(pow((Var*Dispersion*T),Exponent)*T), IntCriteria = 10000*pow(Alpha,-4.534);
	if (Alpha == 2.0)
		IntCriteria = 14;
	if (X < IntCriteria)
	{
		double m = b/2, h = m;
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
			IS = b;
		Result = AlphaStablePDFauxGL(a, b, fa, fb, IS, X, Alpha);
	}
	return Result/Sigma*exp (-Lambda*T);
 }

double LPM::gt_FDMint(double MinAge, double& MaxAge, double& Alpha, double& Tau, double& DP, double Lambda)
{
    double Result=0.0;
	VectorXd ResultArray=VectorXd::Zero(7);

	if ((MaxAge != 0) | (DP != 0))
	{
		double k1 = (MaxAge - MinAge)/2, k2 = (MinAge + MaxAge)/2;
		if (k1<0.1)
		{
			#pragma omp parallel for schedule(dynamic)
			for(int i=0;i<3;i++)			
				ResultArray[i]=k1*Coeff3[i]*AlphaStablePDFInt(k1*Roots3[i]+k2,Alpha,Tau,DP,Lambda);
		}
		else
		{
			#pragma	omp parallel for schedule(dynamic)
			for(int i=0;i<7;i++)	
				ResultArray[i]=k1*Coeff7[i]*AlphaStablePDFInt(k1*Roots7[i]+k2,Alpha,Tau,DP,Lambda);
		}
		Result=ResultArray.sum();
	}
	return Result;
}

 double LPM::gt_PFM(double MinAge, double MaxAge, double Tau, double UZtime)
{
	double Result = 0.0, BeginAge;
	BeginAge = Tau + UZtime;
	if (MinAge < BeginAge && MaxAge >= BeginAge)
		Result = 1.0;
	else
		Result = 0.0;
	return Result;
}

VectorXd LPM::PEM_Int(double Tau, double PEM_Uratio, double PEM_Lratio, double SampleDate)
{
	double PEMhalf1[10],PEMhalf2[10],Lambda2[10],PEMprev[10],EndDates[10];
	double DR, Cin, MinDate, MaxDate;
	double EndAge, BeginAge, nS, nU, nL, nStar;
	double Tstar, pU, pS, pL, TauUpper, TauLower;
	double TauRes, EndDate, PEMnoDecay=0.0, TimeIncrement;
	double Multiplier, MinAge, MaxAge, PEMnd1=0.0, PEMnd2=0.0;
	int i,j, nIters, StepInc, StopCriteria,NameCheck=0,Js[10];
	int TracerCount=0, CheckComplete, TracerNum=0, Cond[10] = {0,0,0,0,0,0,0,0,0,0};

	TracerNum=obj.Tracer.Tracers.size();
	VectorXd Result = VectorXd::Zero(TracerNum);

	if (Tau > 100)
	{
		TimeIncrement = fabs(obj.Tracer.DateRange(1)-obj.Tracer.DateRange(0)); //MinTimeInc(DateRange);
		if (fabs(obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 1)-obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 2)) < TimeIncrement)
		{
			TimeIncrement = fabs(obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 1)-obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 2));
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
		return Result;

	MinDate = SampleDate - BeginAge;
	MaxDate = ReturnEndDate(MinDate,TimeIncrement);
	MinAge = SampleDate - MinDate;
	MaxAge = SampleDate - MaxDate;
	EndDate = MinDate;
	j = obj.Tracer.DateRange.rows() - 1;
	StepInc = 1;
	StopCriteria = 0;
	if (obj.Tracer.DateRange(j) < obj.Tracer.DateRange(0))
	{
		StepInc = -1;
		StopCriteria = j;
		j = 0;
	}
	DR = obj.Tracer.DateRange(j);
	
	TracerParseEM(EndDate,TimeIncrement,TracerCount,SampleDate,EndDates,Js,j,MinDate,DR,StopCriteria,StepInc);

	for (int k = 0;k < TracerNum; k++)
		{	
			Cin = obj.Tracer.TracerInputRange(Js[k],k);
			if(obj.Tracer.Tracers[k]==5)
			{
				if (obj.Tracer.He4.He4SolnRate != 0.0)
					Cin = obj.Tracer.He4.He4SolnRate;
				else
					Cin = obj.Tracer.He4.SedDensity / obj.Tracer.He4.Porosity * (Udecay * obj.Tracer.He4.Uppm + THdecay * obj.Tracer.He4.THppm);
			}
			if (Cin == 0 && j == StopCriteria)
				return Result;
			Multiplier = nS / TauRes * (1 / ((1 / TauRes) + obj.Tracer.Lambda(k)));
			PEMhalf1[k]= exp(-(MinAge) * ((1 / TauRes) + obj.Tracer.Lambda(k)));
			PEMhalf2[k] = exp(-(MaxAge) * ((1 / TauRes) + obj.Tracer.Lambda(k)));

			switch(obj.Tracer.CalcCond[k])
			{
			case 1:
				Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (PEMhalf1[k] - PEMhalf2[k]);
				break;
			case 2:
				Cin*=(1-ReturnLambdaCorrection(MinAge,MaxAge,obj.Lambda2[k]));
				Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (PEMhalf1[k] - PEMhalf2[k]);
				break;
			case 3:
				if(obj.k1==0 && obj.k2==0)
				{
					Multiplier = nS / TauRes * (1 / ((1 / TauRes) + obj.Lambda2[k]));
					PEMnd1 = exp(-(MinAge) / TauRes);
					PEMnd2 = exp(-(MaxAge) / TauRes);
					PEMnoDecay = Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * nS * (PEMnd1 - PEMnd2);
					PEMnd1 = PEMnd2;
					PEMhalf1[k] = exp(-(MinAge) * ((1 / TauRes) + obj.Lambda2[k]));
					PEMhalf2[k] = exp(-(MaxAge) * ((1 / TauRes) + obj.Lambda2[k]));
					Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (PEMhalf1[k] - PEMhalf2[k]);
				}
				else
					Result(k)=Result(obj.k1)/Result(obj.k2); 
				break;
			case 4:
				Result(k) += Cin  * Multiplier * (PEMhalf1[k] - PEMhalf2[k]) * MaxAge;
				break;
			default:
				break;
			}
		}
		
	MinAge = MaxAge;
	for (i = 1; i<= nIters; i++)
	{
		MaxAge = MinAge + i * TimeIncrement;
        //loop through
		CheckComplete=0;
		for(int k=0;k<TracerNum;k++)
		{
			PEMhalf1[k] = PEMhalf2[k];
			DR = obj.Tracer.DateRange(Js[k]);
			MaxDate = EndDates[k]-i * TimeIncrement;
			while (MaxDate < DR && Js[k] != StopCriteria)
			{
				Js[k] -= StepInc;
				DR = obj.Tracer.DateRange(Js[k]);;
			}
			Cin = obj.Tracer.TracerInputRange(Js[k],k);
			if(obj.Tracer.Tracers[k]==5)
			{
				if (obj.Tracer.He4.He4SolnRate != 0.0)
					Cin = obj.Tracer.He4.He4SolnRate;
				else
					Cin = obj.Tracer.He4.SedDensity / obj.Tracer.He4.Porosity * (Udecay * obj.Tracer.He4.Uppm + THdecay * obj.Tracer.He4.THppm);
			}
			if (Cin != 0)
			{
				PEMprev[k]= Result(k);
				Multiplier = nS / TauRes * (1 / ((1 / TauRes) + obj.Tracer.Lambda(k)));
				PEMhalf2[k] = exp(-(MaxAge) * ((1 / TauRes) + obj.Tracer.Lambda(k)));
				switch(obj.Tracer.CalcCond[k])
				{
				case 1:
					Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (PEMhalf1[k] - PEMhalf2[k]);
					break;
				case 2:
					Cin*=(1-ReturnLambdaCorrection(MinAge+(i-1)*TimeIncrement,MaxAge,obj.Lambda2[k]));
					Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (PEMhalf1[k] - PEMhalf2[k]);
					break;
				case 3:
					if(obj.k1==0 && obj.k2==0)
					{
						Multiplier = nS / TauRes * (1 / ((1 / TauRes) + obj.Lambda2[k]));
						PEMnd1 = exp(-(MinAge+(i-1)*TimeIncrement) / TauRes);
						PEMnd2 = exp(-(MaxAge) / TauRes);
						PEMnoDecay = Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * nS * (PEMnd1 - PEMnd2);
						PEMnd1 = PEMnd2;
						PEMhalf2[k] = exp(-(MaxAge) * ((1 / TauRes) + obj.Lambda2[k]));
						Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (PEMhalf1[k] - PEMhalf2[k]);
					}
					else
						Result(k)=Result(obj.k1)/Result(obj.k2);
					break;
				case 4:
					Result(k) += Cin  * Multiplier * (PEMhalf1[k] - PEMhalf2[k]) * MaxAge;
					break;
				}
				if (MaxAge > Tau && Result(k) > 0) //This has to be outer if,then; otherwise Overflow occurs because of division by zero
				{
					if ((Result(k) - PEMprev[k]) / Result(k) < Tol)
						CheckComplete++;
				}
			}
			else //if (Cin == 0 || Js[k] == StopCriteria)
				CheckComplete++;
		}
		if(CheckComplete==TracerCount)
			return Result;
	}
	i++;
	MaxAge = MinAge + i * TimeIncrement;
	if (MaxAge >= EndAge)
	{
		for(int k=0;k<TracerNum;k++)
		{
			PEMhalf1[k] = PEMhalf2[k];
			MaxAge = EndAge;
			DR = obj.Tracer.DateRange(Js[k]);
			MaxDate = EndDates[k]-i * TimeIncrement;
			while (MaxDate < DR && Js[k] != StopCriteria)
			{
				Js[k] = Js[k] - StepInc;
				DR = obj.Tracer.DateRange(Js[k]);;
			}
			Cin = obj.Tracer.TracerInputRange(Js[k],k);
			if(obj.Tracer.Tracers[k]==5)
			{
				if (obj.Tracer.He4.He4SolnRate != 0.0)
					Cin = obj.Tracer.He4.He4SolnRate;
				else
					Cin = obj.Tracer.He4.SedDensity / obj.Tracer.He4.Porosity * (Udecay * obj.Tracer.He4.Uppm + THdecay * obj.Tracer.He4.THppm);
			}
			Multiplier = nS / TauRes * (1 / ((1 / TauRes) + obj.Tracer.Lambda(k)));
			PEMhalf2[k] = exp(-(MaxAge) * ((1 / TauRes) + obj.Tracer.Lambda(k)));
			if (Cin != 0)
			{
				PEMprev[k]= Result(k);
				switch(obj.Tracer.CalcCond[k])
				{
				case 1:
					Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (PEMhalf1[k] - PEMhalf2[k]);
					break;
				case 2:
					Cin*=(1-ReturnLambdaCorrection(MinAge+(i-1)*TimeIncrement,MaxAge,obj.Lambda2[k]));
					Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (PEMhalf1[k] - PEMhalf2[k]);
					break;
				case 3:
					if(obj.k1==0 && obj.k2==0)
					{
						Multiplier = nS / TauRes * (1 / ((1 / TauRes) + obj.Lambda2[k]));
						PEMnd1 = exp(-(MinAge+(i-1)*TimeIncrement) / TauRes);
						PEMnd2 = exp(-(MaxAge) / TauRes);
						PEMnoDecay = Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * nS * (PEMnd1 - PEMnd2);
						PEMnd1 = PEMnd2;
						PEMhalf2[k] = exp(-(MaxAge) * ((1 / TauRes) + obj.Lambda2[k]));
						Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (PEMhalf1[k] - PEMhalf2[k]);
					}
					else
						Result(k)=Result(obj.k1)/Result(obj.k2);
					break;
				case 4:
					Result(k) += Cin  * Multiplier * (PEMhalf1[k] - PEMhalf2[k]) * MaxAge;
					break;
				}
			}
		}
	}
	return Result;
}


VectorXd LPM::GAM_Int(double Tau, double SampleDate, double Alpha)
{
	double GAMhalf1,GAMhalf2, GamFrac,GAMprev[10],EndDates[10];
	double DR, Cin, MinDate, MaxDate, Beta, dGamHalf2=0;
	double EndDate=0, GAMnoDecay=0.0, TimeIncrement;
	double Multiplier, MinAge, MaxAge, GAMnd1=0.0, GAMnd2=0.0;
	int i,j, nIters, StepInc, StopCriteria,NameCheck=0,Js[10];
	int TracerCount=0, CheckComplete, TracerNum=0;

	
	TracerNum=obj.Tracer.Tracers.size();
	VectorXd Result = VectorXd::Zero(TracerNum);

	if (Tau > 100)
	{
		TimeIncrement = fabs(obj.Tracer.DateRange(1)-obj.Tracer.DateRange(0)); 
		if (fabs(obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 1)-obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 2)) < TimeIncrement)
			TimeIncrement = fabs(obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 1)-obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 2));
	}
	else
		TimeIncrement = 1.0 / 12.0;
	Beta = Tau/Alpha;
	boost::math::gamma_distribution<> GAM_dist(Alpha, Beta);	
	MinDate = SampleDate;
	MaxDate = ReturnEndDate(MinDate,TimeIncrement);
	MinAge = SampleDate - MinDate;
	MaxAge = SampleDate - MaxDate;
	EndDate = MinDate;
	nIters = 2000000;
	j = obj.Tracer.DateRange.rows() - 1;
	StepInc = 1;
	StopCriteria = 0;
	if (obj.Tracer.DateRange(j) < obj.Tracer.DateRange(0))
	{
		StepInc = -1;
		StopCriteria = j;
		j = 0;
	}
	DR = obj.Tracer.DateRange(j);
	
	TracerParseEM(EndDate,TimeIncrement,TracerCount,SampleDate,EndDates,Js,j,MinDate,DR,StopCriteria,StepInc);
	GAMhalf1=boost::math::cdf(GAM_dist,MinAge);
	GAMhalf2=boost::math::cdf(GAM_dist,MaxAge);
	GamFrac=GAMhalf2-GAMhalf1;
	for (int k = 0;k < TracerNum; k++)
	{	
		Cin = obj.Tracer.TracerInputRange(Js[k],k);
		if(obj.Tracer.Tracers[k]==5)
		{
			if (obj.Tracer.He4.He4SolnRate != 0.0)
				Cin = obj.Tracer.He4.He4SolnRate;
			else
				Cin = obj.Tracer.He4.SedDensity / obj.Tracer.He4.Porosity * (Udecay * obj.Tracer.He4.Uppm + THdecay * obj.Tracer.He4.THppm);
		}
		if (Cin == 0 && j == StopCriteria)
			return Result;

		switch(obj.Tracer.CalcCond[k])
		{
			case 1:
				Multiplier = ReturnLambdaCorrection(MinAge,MaxAge,obj.Tracer.Lambda(k));
				Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * GamFrac;
				break;
			case 2:
				Cin*=(1-ReturnLambdaCorrection(MinAge,MaxAge,obj.Lambda2[k]));
				Multiplier = ReturnLambdaCorrection(MinAge,MaxAge,obj.Tracer.Lambda(k));
				Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * GamFrac;
				break;
			case 3:
				if(obj.k1==0 && obj.k2==0)
				{
					Multiplier = ReturnLambdaCorrection(MinAge,MaxAge,obj.Lambda2[k]);
						
					GAMnoDecay = Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * GamFrac;
					GAMnd1 = GAMnd2;
					GAMhalf1 = GAMnd1;
					GAMhalf2 = GAMnd2;
					Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (GAMhalf2 - GAMhalf1);
				}
				else 
					Result(k)=Result(obj.k1)/Result(obj.k2);
				break;
			case 4:
				Result(k) += Cin * GamFrac * MaxAge;
				break;
			default:
				break;
		}
	}
	MinAge = MaxAge;
	for (i = 1; i<= nIters; i++)
	{
		MaxAge = MinAge + i * TimeIncrement;
		GAMhalf1=GAMhalf2;
		GAMhalf2=boost::math::cdf(GAM_dist,MaxAge);
		GamFrac=GAMhalf2-GAMhalf1;
        //loop through
		CheckComplete=0;
		for(int k=0;k<TracerNum;k++)
		{
			DR = obj.Tracer.DateRange(Js[k]);
			MaxDate = EndDates[k]-i * TimeIncrement;
			while (MaxDate < DR && Js[k] != StopCriteria)
			{
				Js[k] -= StepInc;
				DR = obj.Tracer.DateRange(Js[k]);
			}
				
			Cin = obj.Tracer.TracerInputRange(Js[k],k);
			if(obj.Tracer.Tracers[k]==5)
			{
				if (obj.Tracer.He4.He4SolnRate != 0.0)
					Cin = obj.Tracer.He4.He4SolnRate;
				else
					Cin = obj.Tracer.He4.SedDensity / obj.Tracer.He4.Porosity * (Udecay * obj.Tracer.He4.Uppm + THdecay * obj.Tracer.He4.THppm);
			}

			if (Cin != 0)
			{
				GAMprev[k]= Result(k);
				switch(obj.Tracer.CalcCond[k])
				{
					case 1:
						Multiplier = ReturnLambdaCorrection(MinAge+(i-1)*TimeIncrement,MaxAge,obj.Tracer.Lambda(k));
						Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * GamFrac;
						break;
					case 2:
						Cin*=(1-ReturnLambdaCorrection(MinAge+(i-1)*TimeIncrement,MaxAge,obj.Lambda2[k]));
						Multiplier = ReturnLambdaCorrection(MinAge+(i-1)*TimeIncrement,MaxAge,obj.Tracer.Lambda(k));
						Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * GamFrac;
						break;
					case 3:
						if(obj.k1==0 && obj.k2==0)
						{
							Multiplier = ReturnLambdaCorrection(MinAge+(i-1)*TimeIncrement,MaxAge,obj.Lambda2[k]);
							GAMnd1 = boost::math::cdf(GAM_dist,MinAge+(i-1)*TimeIncrement);
							GAMnd2 = GAMhalf2; //I think this is supposed to be GAMhalf2 not dGamHalf2
							GAMnoDecay = obj.Tracer.TracerInputRange(Js[k],k) * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * (GAMnd2 - GAMnd1);
							GAMnd1 = GAMnd2;
							GAMhalf1 = GAMnd1;
							GAMhalf2 = GAMnd2;
							Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (GAMhalf2 - GAMhalf1);
						}
						else
							Result(k)=Result(obj.k1)/Result(obj.k2);
						break;
					case 4:
						Result(k) += Cin * GamFrac * MaxAge;
						break;
					default:
						break;
				}
				if (MaxAge > Tau && Result(k) > 0) //This has to be outer if,then; otherwise Overflow occurs because of division by zero
				{
					if ((Result(k) - GAMprev[k]) / Result(k) < Tol)
						CheckComplete++;					
				}
			}
			else //if (Cin == 0 || Js[k] == StopCriteria)
				CheckComplete++;
		}
		if(CheckComplete==TracerCount)
			return Result;
	}
	return Result;
}

VectorXd LPM::EMM_Int(double Tau, double SampleDate)
{
	double EMMhalf1[10],EMMhalf2[10],Lambda2[10],EMMprev[10],EndDates[10];
	double DR, Cin, MinDate, MaxDate;
	double n, EndDate, EMMnoDecay=0.0, TimeIncrement;
	double Multiplier, MinAge, MaxAge, EMMnd1=0.0, EMMnd2=0.0;
	int i,j, nIters, StepInc, StopCriteria,NameCheck=0,Js[10];
	int TracerCount=0, CheckComplete, TracerNum=0;

	TracerNum=obj.Tracer.Tracers.size();
	VectorXd Result = VectorXd::Zero(TracerNum);

	if (Tau > 100)
	{
		TimeIncrement = fabs(obj.Tracer.DateRange(1)-obj.Tracer.DateRange(0)); //MinTimeInc(obj.Tracer.DateRange);
		if (fabs(obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 1)-obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 2)) < TimeIncrement)
			TimeIncrement = fabs(obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 1)-obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 2));
	}
	else
		TimeIncrement = 1.0 / 12.0;
	n = 1;
	MinDate = SampleDate; //Age at Z star
	MaxDate = ReturnEndDate(MinDate,TimeIncrement);
	MinAge = SampleDate - MinDate;
	MaxAge = SampleDate - MaxDate;
	EndDate = MinDate;
	nIters = 2000000;
	j = obj.Tracer.DateRange.rows() - 1;
	StepInc = 1;
	StopCriteria = 0;
	if (obj.Tracer.DateRange(j) < obj.Tracer.DateRange(0))
	{
		StepInc = -1;
		StopCriteria = j;
		j = 0;
	}
	DR = obj.Tracer.DateRange(j);
	
	TracerParseEM(EndDate,TimeIncrement,TracerCount,SampleDate,EndDates,Js,j,MinDate,DR,StopCriteria, StepInc);
	
	for (int k = 0;k < TracerNum; k++)
	{
		Cin = obj.Tracer.TracerInputRange(Js[k],k);
		if(obj.Tracer.Tracers[k]==5)
		{
			if (obj.Tracer.He4.He4SolnRate != 0.0)
				Cin = obj.Tracer.He4.He4SolnRate;
			else
				Cin = obj.Tracer.He4.SedDensity / obj.Tracer.He4.Porosity * (Udecay * obj.Tracer.He4.Uppm + THdecay * obj.Tracer.He4.THppm);
		}
		if (Cin == 0 && j == StopCriteria)
			return Result;

		switch(obj.Tracer.CalcCond[k])
		{
		case 1:
			Multiplier = n / Tau * (1 / ((1 / Tau) + obj.Tracer.Lambda(k)));
			EMMhalf1[k]= exp(-(MinAge) * ((1 / Tau) + obj.Tracer.Lambda(k)));
			EMMhalf2[k] = exp(-(MaxAge) * ((1 / Tau) + obj.Tracer.Lambda(k)));
			Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (EMMhalf1[k] - EMMhalf2[k]);
			break;
		case 2:
			Cin=(1-ReturnLambdaCorrection(MinAge,MaxAge,obj.Lambda2[k]))*Cin;
			Multiplier = n / Tau * (1 / ((1 / Tau) + obj.Tracer.Lambda(k)));
			EMMhalf1[k]= exp(-(MinAge) * ((1 / Tau) + obj.Tracer.Lambda(k)));
			EMMhalf2[k] = exp(-(MaxAge) * ((1 / Tau) + obj.Tracer.Lambda(k)));
			Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (EMMhalf1[k] - EMMhalf2[k]);
			break;
		case 3:
			if(obj.k1==0 && obj.k2==0)
			{
				Multiplier = n / Tau * (1 / ((1 / Tau) + obj.Lambda2[k]));
				EMMnd1 = exp(-(MinAge) / Tau);
				EMMnd2 = exp(-(MaxAge) / Tau);
				EMMnoDecay = obj.Tracer.TracerInputRange(Js[k],k) * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * n * (EMMnd1 - EMMnd2);
				EMMnd1 = EMMnd2;
				EMMhalf1[k] = exp(-(MinAge) * ((1 / Tau) + obj.Lambda2[k]));
				EMMhalf2[k] = exp(-(MaxAge) * ((1 / Tau) + obj.Lambda2[k]));
				Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (EMMhalf1[k] - EMMhalf2[k]);
			}
			else
				Result(k)=Result(obj.k1)/Result(obj.k2); 
			break;
		case 4:
			Multiplier = n / Tau * (1 / ((1 / Tau) + obj.Tracer.Lambda(k)));
			EMMhalf1[k]= exp(-(MinAge) * ((1 / Tau) + obj.Tracer.Lambda(k)));
			EMMhalf2[k] = exp(-(MaxAge) * ((1 / Tau) + obj.Tracer.Lambda(k)));
			Result(k) += Cin * Multiplier * (EMMhalf1[k] - EMMhalf2[k]) * MaxAge;
			break;
		}
	}
	MinAge = MaxAge;
	for (i = 1; i<= nIters; i++)
	{
		MaxAge = MinAge + i * TimeIncrement;
        //loop through
		CheckComplete=0;
		for(int k=0;k<TracerNum;k++)
		{
			EMMhalf1[k] = EMMhalf2[k];
			DR = obj.Tracer.DateRange(Js[k]);
			MaxDate = EndDates[k]-i * TimeIncrement;
			while (MaxDate < DR && Js[k] != StopCriteria)
			{
				Js[k] -= StepInc;
				DR = obj.Tracer.DateRange(Js[k]);
			}
			Cin = obj.Tracer.TracerInputRange(Js[k],k);
			if(obj.Tracer.Tracers[k]==5)
			{
				if (obj.Tracer.He4.He4SolnRate != 0.0)
					Cin = obj.Tracer.He4.He4SolnRate;
				else
					Cin = obj.Tracer.He4.SedDensity / obj.Tracer.He4.Porosity * (Udecay * obj.Tracer.He4.Uppm + THdecay * obj.Tracer.He4.THppm);
			}
			if (Cin != 0)
			{
				EMMprev[k]= Result(k);
				switch(obj.Tracer.CalcCond[k])
				{
				case 1:
					Multiplier = n / Tau * (1 / ((1 / Tau) + obj.Tracer.Lambda(k)));
					EMMhalf2[k] = exp(-(MaxAge) * ((1 / Tau) + obj.Tracer.Lambda(k)));
					Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (EMMhalf1[k] - EMMhalf2[k]);
					break;
				case 2: //3He(trit)
					Cin*=(1-ReturnLambdaCorrection(MinAge+(i-1)*TimeIncrement,MaxAge,obj.Lambda2[k]));
					Multiplier = n / Tau * (1 / ((1 / Tau) + obj.Tracer.Lambda(k)));
					EMMhalf2[k] = exp(-(MaxAge) * ((1 / Tau) + obj.Tracer.Lambda(k)));
					Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (EMMhalf1[k] - EMMhalf2[k]);
					break;
				case 3: //3H/3Ho
					if(obj.k1==0 && obj.k2==0)
					{
						Multiplier = n / Tau * (1 / ((1 / Tau) + obj.Lambda2[k]));
						EMMnd1 = exp(-(MinAge+(i-1)*TimeIncrement) / Tau);
						EMMnd2 = exp(-(MaxAge) / Tau);
						EMMnoDecay = obj.Tracer.TracerInputRange(Js[k],k) * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * n * (EMMnd1 - EMMnd2);
						EMMnd1 = EMMnd2;
						EMMhalf2[k] = exp(-(MaxAge) * ((1 / Tau) + obj.Lambda2[k]));
						Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (EMMhalf1[k] - EMMhalf2[k]);	
					}
					else
						Result(k)=Result(obj.k1)/Result(obj.k2);
					break;
				case 4:
					Multiplier = n / Tau * (1 / ((1 / Tau) + obj.Tracer.Lambda(k)));
					EMMhalf2[k] = exp(-(MaxAge) * ((1 / Tau) + obj.Tracer.Lambda(k)));
					Result(k) += Cin * Multiplier * (EMMhalf1[k] - EMMhalf2[k]) * MaxAge;
					break;
				}

				//check only for first tracer \/
				if (MaxAge > Tau && Result(k) > 0) //This has to be outer if,then; otherwise Overflow occurs because of division by zero
				{
					if ((Result(k) - EMMprev[k]) / Result(k) < Tol)
						CheckComplete++;					
				}
			}
			else //if (Cin == 0 || Js[k] == StopCriteria)
				CheckComplete++;
		}
		if(CheckComplete==TracerCount)
			return Result;
	}
	return Result;
}

VectorXd LPM::EPM_Int(double Tau, double SampleDate, double EPMratio)
{

	double EPMhalf1[10],EPMhalf2[10],Lambda2[10],EPMprev[10],EndDates[10];
	double DR, Cin, MinDate, MaxDate;
	double n, EndDate, EPMnoDecay=0.0, TimeIncrement;
	double Multiplier, MinAge, MaxAge, EPMnd1=0.0, EPMnd2=0.0;
	int i,j, nIters, StepInc, StopCriteria,NameCheck=0,Js[10];
	int TracerCount=0, CheckComplete, TracerNum=0;

	TracerNum=obj.Tracer.Tracers.size();
	VectorXd Result = VectorXd::Zero(TracerNum);

	if (Tau > 100)
	{
		TimeIncrement = fabs(obj.Tracer.DateRange(1)-obj.Tracer.DateRange(0)); //MinTimeInc(obj.Tracer.DateRange);
		if (fabs(obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 1)-obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 2)) < TimeIncrement)
			TimeIncrement = fabs(obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 1)-obj.Tracer.DateRange(obj.Tracer.DateRange.rows() - 2));
	}
	else
		TimeIncrement = 1.0 / 12.0;
	
	n = EPMratio + 1;
	MinAge = Tau*(1.0 - (1.0 / n)); //For EPM
	MinDate = SampleDate - MinAge; //Age at Z star
	MaxDate = ReturnEndDate(MinDate,TimeIncrement);
	MinAge = SampleDate - MinDate;
	MaxAge = SampleDate - MaxDate;
	EndDate = MinDate;

	nIters = 2000000;
	j = obj.Tracer.DateRange.rows() - 1;
	StepInc = 1;
	StopCriteria = 0;
	if (obj.Tracer.DateRange(j) < obj.Tracer.DateRange(0))
	{
		StepInc = -1;
		StopCriteria = j;
		j = 0;
	}
	DR = obj.Tracer.DateRange(j);
	
	TracerParseEM(EndDate,TimeIncrement,TracerCount,SampleDate,EndDates,Js,j,MinDate,DR,StopCriteria, StepInc);
	
	for (int k = 0;k < TracerNum; k++)
	{		
		Cin = obj.Tracer.TracerInputRange(Js[k],k);
		if(obj.Tracer.Tracers[k]==5)
		{
			if (obj.Tracer.He4.He4SolnRate != 0.0)
				Cin = obj.Tracer.He4.He4SolnRate;
			else
				Cin = obj.Tracer.He4.SedDensity / obj.Tracer.He4.Porosity * (Udecay * obj.Tracer.He4.Uppm + THdecay * obj.Tracer.He4.THppm);
		}
		if (Cin == 0 && j == StopCriteria)
			return Result;

		switch(obj.Tracer.CalcCond[k])
		{
			case 1:
				Multiplier = n / Tau * (1 / (n / Tau + obj.Tracer.Lambda(k)));
				EPMhalf1[k]= exp(-MinAge * (n / Tau + obj.Tracer.Lambda(k)) + n - 1.0);
				EPMhalf2[k] = exp(-MaxAge * (n / Tau + obj.Tracer.Lambda(k)) + n - 1.0);
				Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (EPMhalf1[k] - EPMhalf2[k]);
				break;
			case 2:
				Cin*=(1-ReturnLambdaCorrection(MinAge,MaxAge,obj.Lambda2[k]));
				Multiplier = n / Tau * (1 / (n / Tau + obj.Tracer.Lambda(k)));
				EPMhalf1[k]= exp(-MinAge * (n / Tau + obj.Tracer.Lambda(k)) + n - 1.0);
				EPMhalf2[k] = exp(-MaxAge * (n / Tau + obj.Tracer.Lambda(k)) + n - 1.0);
				Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (EPMhalf1[k] - EPMhalf2[k]);
				break;
			case 3:

				if(obj.k1==0 && obj.k2==0)
				{
					Multiplier = n / Tau * (1 / (n / Tau + obj.Lambda2[k]));
					EPMnd1 = exp(-MinAge * n / Tau + n - 1);
					EPMnd2 = exp(-MaxAge * n / Tau + n - 1);
					EPMnoDecay = obj.Tracer.TracerInputRange(Js[k],k) * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * (EPMnd1 - EPMnd2);
					EPMnd1 = EPMnd2;
					EPMhalf1[k] = exp(-MinAge * (n / Tau + obj.Lambda2[k]) + n - 1);
					EPMhalf2[k] = exp(-MaxAge * (n / Tau + obj.Lambda2[k]) + n - 1);
					Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (EPMhalf1[k] - EPMhalf2[k]);
				}

				else if(obj.k1!=0 || obj.k2!=0)
					Result(k)=Result(obj.k1)/Result(obj.k2); 
				break;
			case 4:
				Multiplier = n / Tau * (1 / (n / Tau + obj.Tracer.Lambda(k)));
				EPMhalf1[k]= exp(-MinAge * (n / Tau + obj.Tracer.Lambda(k)) + n - 1.0);
				EPMhalf2[k] = exp(-MaxAge * (n / Tau + obj.Tracer.Lambda(k)) + n - 1.0);
				Result(k) += Cin * Multiplier * (EPMhalf1[k] - EPMhalf2[k]) * MaxAge;
				break;
		}
	}
	MinAge = MaxAge;
	for (i = 1; i<= nIters; i++)
	{
		MaxAge = MinAge + i * TimeIncrement;
        //loop through
		CheckComplete=0;
		for(int k=0;k<TracerNum;k++)
		{
			EPMhalf1[k] = EPMhalf2[k];
			DR = obj.Tracer.DateRange(Js[k]);
			MaxDate = EndDates[k]-i * TimeIncrement;
			while (MaxDate < DR && Js[k] != StopCriteria)
			{
				Js[k] -= StepInc;
				DR = obj.Tracer.DateRange(Js[k]);
			}
				
			Cin = obj.Tracer.TracerInputRange(Js[k],k);
			if(obj.Tracer.Tracers[k]==5)
			{
				if (obj.Tracer.He4.He4SolnRate != 0.0)
					Cin = obj.Tracer.He4.He4SolnRate;
				else
					Cin = obj.Tracer.He4.SedDensity / obj.Tracer.He4.Porosity * (Udecay * obj.Tracer.He4.Uppm + THdecay * obj.Tracer.He4.THppm);
			}
			if (Cin != 0)
			{
				EPMprev[k]= Result(k);
				switch(obj.Tracer.CalcCond[k])
					{
					case 1:
						Multiplier = n / Tau * (1 / (n / Tau + obj.Tracer.Lambda(k)));
						EPMhalf2[k] = exp(-MaxAge * (n / Tau + obj.Tracer.Lambda(k)) + n - 1.0);
						Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (EPMhalf1[k] - EPMhalf2[k]);
						break;
					case 2:
						Cin*=(1-ReturnLambdaCorrection(MinAge+(i-1)*TimeIncrement,MaxAge,obj.Lambda2[k]));
						Multiplier = n / Tau * (1 / (n / Tau + obj.Tracer.Lambda(k)));
						EPMhalf2[k] = exp(-MaxAge * (n / Tau + obj.Tracer.Lambda(k)) + n - 1.0);
						Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (EPMhalf1[k] - EPMhalf2[k]);
						break;
							
					case 3:
						if(obj.k1==0 && obj.k2==0) //2
						{	
							Multiplier = n / Tau * (1 / (n / Tau + obj.Lambda2[k]));
							EPMnd1 = exp(-(MinAge+(i-1)*TimeIncrement) * n / Tau + n - 1);
							EPMnd2 = exp(-MaxAge * n / Tau + n - 1);
							EPMnoDecay = obj.Tracer.TracerInputRange(Js[k],k) * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * (EPMnd1 - EPMnd2);
							EPMnd1 = EPMnd2;
							EPMhalf1[k] = exp(-(MinAge+(i-1)*TimeIncrement) * (n / Tau + obj.Lambda2[k]) + n - 1);
							EPMhalf2[k] = exp(-MaxAge * (n / Tau + obj.Lambda2[k]) + n - 1);
							Result(k) += Cin * exp(-obj.Lambda2[k] * obj.Tracer.UZtime(k)) * Multiplier * (EPMhalf1[k] - EPMhalf2[k]);
						}
						else if(obj.k1!=0 || obj.k2!=0) //2
							Result(k)=Result(obj.k1)/Result(obj.k2);
						break;
					case 4:
						Multiplier = n / Tau * (1 / (n / Tau + obj.Tracer.Lambda(k)));
						EPMhalf2[k] = exp(-MaxAge * (n / Tau + obj.Tracer.Lambda(k)) + n - 1.0);
						Result(k) += Cin * Multiplier * (EPMhalf1[k] - EPMhalf2[k]) * MaxAge;
						break;
					}

				//check only for first tracer \/
				if (MaxAge > Tau && Result(k) > 0) //This has to be outer if,then; otherwise Overflow occurs because of division by zero
				{
					if ((Result(k) - EPMprev[k]) / Result(k) < Tol)
						CheckComplete++;					
				}
			}
			else //if (Cin == 0 || Js[k] == StopCriteria)
				CheckComplete++;
		}
		if(CheckComplete==TracerCount)
			return Result;
	}
	return Result;
}

VectorXd LPM::LPM_TracerOutput(double MeanAge, double ModelParm1, double ModelParm2, double Fraction, double MeanAge_2, double ModelParm1_2, double ModelParm2_2, 
						double DIC_1, double DIC_2) 
{

	double sampleDate, test;
	__int64 n, TracerNum=0, BinaryNum=0, SingleNum=0, NumActiveVals=0, ActiveCount=0;
		
	TracerNum=obj.Tracer.Tracers.size();
	n=obj.Sample.SampleDates.size();
	
	VectorXd CTwo, TracerOutput;
	NumActiveVals=obj.Sample.ActiveVals.sum();
	VectorXd Result = VectorXd(NumActiveVals);
	for(int i=0;i<n;i++)
	{
		sampleDate=obj.Sample.SampleDates(i);
		SingleNum = obj.Model.ModelNum;
		if(obj.Model.ModelNum>7)
		{ 			
			BinaryNum=obj.Model.ModelNum%10;
			SingleNum = int(obj.Model.ModelNum/10);
			if(obj.Tracer.TracerComp_2.isValid==true)
				CTwo=obj.Tracer.TracerComp_2.Val.block(i,0,i,TracerNum-1);
			else
			{
				switch(BinaryNum)
				{
				case 1:
					CTwo=PFM_Int(MeanAge_2,sampleDate); //1				
				break;
	
				case 2:
					CTwo=EMM_Int(MeanAge_2,sampleDate); //2
				break;

				case 3:
					CTwo=EPM_Int(MeanAge_2,sampleDate,ModelParm1_2); //3
				break;

				case 4:
					CTwo=GAM_Int(MeanAge_2,sampleDate,ModelParm1_2);	//4
				break;

				case 5:
					CTwo=DM_Internal(MeanAge_2,sampleDate,ModelParm1_2);	//5
				break;

				case 6:
					CTwo=PEM_Int(MeanAge_2,ModelParm1_2,ModelParm2_2,sampleDate); //6
				break;
	
				case 7:
					CTwo=FDM_Internal(MeanAge_2,ModelParm2_2,sampleDate,ModelParm1_2);	//7
				break;
				}
			}
		}
		switch(SingleNum)
		{
			case 1:
				TracerOutput=PFM_Int(MeanAge,sampleDate); //1				
			break;
	
			case 2:
				TracerOutput=EMM_Int(MeanAge,sampleDate); //2
			break;

			case 3:
				TracerOutput=EPM_Int(MeanAge,sampleDate,ModelParm1); //3
			break;

			case 4:
				TracerOutput=GAM_Int(MeanAge,sampleDate,ModelParm1);	//4
			break;

			case 5:
				TracerOutput=DM_Internal(MeanAge,sampleDate,ModelParm1);	//5
			break;

			case 6:
				TracerOutput=PEM_Int(MeanAge,ModelParm1,ModelParm2,sampleDate); //6
			break;
	
			case 7:
				TracerOutput=FDM_Internal(MeanAge,ModelParm2,sampleDate,ModelParm1);	//7
			break;

			default:
				
				break;
		}
		
		if(obj.Model.ModelNum>7)
		{
			for(int j=0;j<TracerNum;j++)
			{
				if(obj.Tracer.Tracers[j]==6 && (DIC_1>0 && DIC_2>0) && obj.Sample.ActiveVals(i,j)>0)
				{
					test=Result(ActiveCount)=(TracerOutput(j)*Fraction*DIC_1 + (1-Fraction)*CTwo(j)*DIC_2)/(Fraction*DIC_1+(1-Fraction)*DIC_2);
					ActiveCount++;
				}
				else if(obj.Sample.ActiveVals(i,j)>0)
				{
					test=Result(ActiveCount)=TracerOutput(j)*Fraction + (1-Fraction)*CTwo(j);
					ActiveCount++;
				}
			}
		}
		else
		{
			for(int j=0;j<TracerNum;j++)
			{
				if(obj.Sample.ActiveVals(i,j)>0)
				{
					test=Result(ActiveCount)=TracerOutput(j);
					ActiveCount++;
				}
			}
		}
	}
	return Result;
}


VectorXd LPM::d_dx_LPM_Model(double MeanAge, double ModelParm1, double ModelParm2, double Fraction, double MeanAge_2,
							double ModelParm1_2, double ModelParm2_2, double DIC_1, double DIC_2, int iParm)
{
	VectorXd xUZtime, return1, return2, return3, return4, Result;
	bool IsZero;
	
	double Delta, ScaleFact;
	__int64 n, k;

	n=obj.Sample.SampleDates.size();
	Result.resize(obj.Sample.ActiveVals.sum());
	ScaleFact = 1.0;
	k = 1;
	do
	{
		switch (iParm)
		{
		case 1: //uzTime
			Delta = 0.01*ScaleFact;
			
			for (int i = 0; i < obj.Tracer.UZtime.size(); i++)   // -2*delta
			{
				if (obj.Tracer.UZtimeCond[i] == 1)
					obj.Tracer.UZtime(i) -= (2 * Delta);
			}

			return1 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2, Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);

			for (int i = 0; i < obj.Tracer.UZtime.size(); i++)  //-delta
			{
				if (obj.Tracer.UZtimeCond[i] == 1)
					obj.Tracer.UZtime(i) += Delta;
			}

			return2 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2, Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);

			for (int i = 0; i < obj.Tracer.UZtime.size(); i++)  //+delta
			{
				if (obj.Tracer.UZtimeCond[i] == 1)
					obj.Tracer.UZtime(i) += 2 * Delta;
			}

			return3 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2, Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);

			for (int i = 0; i < obj.Tracer.UZtime.size(); i++)  //+2*delta
			{
				if (obj.Tracer.UZtimeCond[i] == 1)
					obj.Tracer.UZtime(i) += Delta;
			}

			return4 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2, Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);

			for (int i = 0; i < obj.Tracer.UZtime.size(); i++)   //set back to original values
			{
				if (obj.Tracer.UZtimeCond[i] == 1)
					obj.Tracer.UZtime(i) -= 2 * Delta;
			}

			break;
		case 2:  //MeanAge derivative
			(MeanAge < 5) ? Delta = 0.01*ScaleFact : Delta = 0.5*ScaleFact;
			//Delta = 0.01*ScaleFact;
			return1 = LPM_TracerOutput(MeanAge - (2 * Delta), ModelParm1, ModelParm2,
				Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			return2 = LPM_TracerOutput(MeanAge - Delta, ModelParm1, ModelParm2,
				Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			return3 = LPM_TracerOutput(MeanAge + Delta, ModelParm1, ModelParm2,
				Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			return4 = LPM_TracerOutput(MeanAge + (2 * Delta), ModelParm1, ModelParm2,
				Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			break;
		case 3: //Model Paramater 1 derivative
			Delta = 0.0001*ScaleFact;
			return1 = LPM_TracerOutput(MeanAge, ModelParm1 - (2 * Delta), ModelParm2,
				Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			return2 = LPM_TracerOutput(MeanAge, ModelParm1 - Delta, ModelParm2,
				Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			return3 = LPM_TracerOutput(MeanAge, ModelParm1 + Delta, ModelParm2,
				Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			return4 = LPM_TracerOutput(MeanAge, ModelParm1 + (2 * Delta), ModelParm2,
				Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			break;
		case 4: //Model Parameter 2 derivative
			Delta = 0.0001*ScaleFact;
			return1 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2 - (2 * Delta),
				Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			return2 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2 - Delta,
				Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			return3 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2 + Delta,
				Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			return4 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2 + (2 * Delta),
				Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			break;
		case 5: //Mixing fraction 1 derivative
			Delta = 0.0001*ScaleFact;
			return1 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2,
				Fraction - (2 * Delta), MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			return2 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2,
				Fraction - Delta, MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			return3 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2,
				Fraction + Delta, MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			return4 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2,
				Fraction + (2 * Delta), MeanAge_2, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			break;
		case 6: //Mean Age component 2 derivative
			Delta = 0.01*ScaleFact;
			return1 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2,
				Fraction, MeanAge_2 - (2 * Delta), ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			return2 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2,
				Fraction, MeanAge_2 - Delta, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			return3 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2,
				Fraction, MeanAge_2 + Delta, ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			return4 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2,
				Fraction, MeanAge_2 + (2 * Delta), ModelParm1_2, ModelParm2_2, DIC_1, DIC_2);
			break;
		case 7: //Model parameter 1 component 2 derivative
			Delta = 0.0001*ScaleFact;
			return1 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2,
				Fraction, MeanAge_2, ModelParm1_2 - (2 * Delta), ModelParm2_2, DIC_1, DIC_2);
			return2 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2,
				Fraction, MeanAge_2, ModelParm1_2 - Delta, ModelParm2_2, DIC_1, DIC_2);
			return3 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2,
				Fraction, MeanAge_2, ModelParm1_2 + Delta, ModelParm2_2, DIC_1, DIC_2);
			return4 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2,
				Fraction, MeanAge_2, ModelParm1_2 + (2 * Delta), ModelParm2_2, DIC_1, DIC_2);
			break;
		case 8: //Model parameter 2 component 2 derivative
			Delta = 0.0001*ScaleFact;
			return1 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2,
				Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2 - (2 * Delta), DIC_1, DIC_2);
			return2 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2,
				Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2 - Delta, DIC_1, DIC_2);
			return3 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2,
				Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2 + Delta, DIC_1, DIC_2);
			return4 = LPM_TracerOutput(MeanAge, ModelParm1, ModelParm2,
				Fraction, MeanAge_2, ModelParm1_2, ModelParm2_2 + (2 * Delta), DIC_1, DIC_2);
			break;
		default:
			return Result;
			break;
		};
		double test;
		for (int i = 0; i < return4.size(); i++)
		{
			IsZero = false;
			test = Result(i) = (return1(i) - 8 * return2(i) + 8 * return3(i) - return4(i)) / (12 * Delta);
			if (Result(i) == 0)
			{
				IsZero = true;
			}
		}
		if(IsZero==true)
			ScaleFact = pow(10,k);
		k++;
	}while (IsZero && k<4);
	return Result;
}

void LPM::TracerParseEM(double EndDate, double TimeIncrement, int &TracerCount, double SampleDate,double EndDates[],  //!DM, !FDM
	int Js[], int j, double &MinDate,double &DR,int StopCriteria, int StepInc)
{
	for(int k=0;k<obj.Tracer.Tracers.size();k++)
	{
		if(obj.Tracer.Tracers[k]>=0)
		{
			TracerCount++;
			MinDate = EndDate-obj.Tracer.UZtime(k);
			EndDates[k] = ReturnEndDate(MinDate,TimeIncrement);
			Js[k] = j;
			DR = obj.Tracer.DateRange(Js[k]);
			while ((EndDates[k]) < DR && Js[k] != StopCriteria)
			{
				Js[k] = Js[k] - StepInc;
				DR = obj.Tracer.DateRange(Js[k]);
			}
		}
	}
}

std::vector<double> ThomasAlgorithimSingleValue(
	const double lower_diagonal, const double middle_diagonal,
	const double upper_diagonal,
	const std::vector<double> rhs_vector) {
	const size_t matrix_size = rhs_vector.size();
	assert(matrix_size > 0);

	std::vector<double> rhs_vector_prime = rhs_vector;
	std::vector<double> middle_diagonal_prime(matrix_size);
	middle_diagonal_prime[0] = middle_diagonal;
	double m;
	for (size_t i = 1; i < matrix_size; ++i) {
		m = lower_diagonal / middle_diagonal_prime[i - 1];
		middle_diagonal_prime[i] = middle_diagonal - m * upper_diagonal;
		rhs_vector_prime[i] = rhs_vector_prime[i] - m * rhs_vector_prime[i - 1];
	}

	std::vector<double> solution_vector(matrix_size);
	solution_vector.back() = (rhs_vector_prime.back() /
		middle_diagonal_prime.back());
	for (size_t i = matrix_size - 1; i > 0; i--) {
		solution_vector[i - 1] = (rhs_vector_prime[i - 1] - upper_diagonal *
			solution_vector[i]) /
			middle_diagonal_prime[i - 1];
	}

	return solution_vector;
}

LPXLOPER12 FullyImplicitAtDepthXL(
	const double delta_time, const double TotSimTime, const double delta_depth, const double max_depth,
	const double effective_diffusion, const double effective_velocity, const double decay_rate, const double requested_depth,
	LPXLOPER12 &surface_tracer_concentrations, bool IsDescending) {

	const size_t NumTimeSteps = TotSimTime / delta_time;
	const size_t NumDepthSteps = max_depth / delta_depth + 1; 
	
	// Construct factors in the iterative equation obtained from the finite
	// difference method. These parameters are described in the
	// solver_method.tex file.
	const double alpha = (effective_diffusion * delta_time) /
		(std::pow(delta_depth, 2));
	const double beta = (effective_velocity * delta_time) / (2 * delta_depth);
	// Construct the diagonal entries of the tridiagonal matrix.
	const double current_time_lower_diagonal = -alpha - beta;
	const double current_time_middle_diagonal = 1 + decay_rate + 2 * alpha;
	const double current_time_upper_diagonal = -alpha + beta;
	LPXLOPER12 pT;
	
	// Initialize the previous time solution with boundary condition at t = 0
	std::vector<double> previous_time_solution(NumDepthSteps, 0.);
	if (IsDescending)
		pT = surface_tracer_concentrations->val.array.lparray + surface_tracer_concentrations->val.array.rows - 1;
	else
		pT = surface_tracer_concentrations->val.array.lparray;
	previous_time_solution[0] = pT->val.num;
//	previous_time_solution[0] = surface_tracer_concentrations[0];
	// Find the depth step closest to the requested depth
	size_t requested_depth_step = round(requested_depth / delta_depth);
	//std::vector<double> solution_at_requested_depth(time_steps + 1);
	// Add initial value to solution_at_requested_depth;
	//solution_at_requested_depth[0] = previous_time_solution[requested_depth_step];
	double XL_size = (NumTimeSteps);
	XLOPER12 *solution = (XLOPER12 *)malloc(XL_size * sizeof(XLOPER12));
	std::vector<double> previous_time_vector(NumDepthSteps-1);
	for (size_t time_step = 1; time_step < NumTimeSteps + 1; ++time_step) {
		// Calculate the RHS vector from the previous time step. This is simply the
		// previous time step solution, plus some boundary offsets.
		std::copy(std::next(previous_time_solution.begin(), 1),
			std::prev(previous_time_solution.end(), 1),
			previous_time_vector.begin());
		// Now add boundary condition offsets.
		// Note, the boundary offset at max depth is 0 because boundary at
		// max_depth is 0.
		if (time_step < surface_tracer_concentrations->val.array.rows) {
			if (IsDescending)
				pT = surface_tracer_concentrations->val.array.lparray + surface_tracer_concentrations->val.array.rows - time_step - 1;
			else
				pT = surface_tracer_concentrations->val.array.lparray + time_step;
		}
		previous_time_vector[0] -= (current_time_lower_diagonal *
			pT->val.num);
		//previous_time_vector[0] -= (current_time_lower_diagonal *
		//	surface_tracer_concentrations[time_step]);

		// Calculate solution for this time step
		std::vector<double> current_time_solution =
			ThomasAlgorithimSingleValue(
				current_time_lower_diagonal, current_time_middle_diagonal,
				current_time_upper_diagonal, previous_time_vector);
		// Insert the solution into the previous_time_solution
		const auto insert_start = std::next(previous_time_solution.begin(), 1);
		std::copy(current_time_solution.begin(), current_time_solution.end(),
			insert_start);
		// Set surface boundary value of previous_time_solution
		previous_time_solution[0] = pT->val.num;
		// Add the relevant value to the wanted_depth_solution
		if (IsDescending) {
			solution[NumTimeSteps-time_step].xltype = xltypeNum;
			solution[NumTimeSteps-time_step].val.num = previous_time_solution[requested_depth_step];
		}
		else {
			solution[time_step-1].xltype = xltypeNum;
			solution[time_step-1].val.num = previous_time_solution[requested_depth_step];
		}
		//solution_at_requested_depth[time_step] =
		//	previous_time_solution[requested_depth_step];
	}
	// Create an array of pointers to XLOPER12 values.
	LPXLOPER12 pArray = (LPXLOPER12)malloc(XL_size * sizeof(LPXLOPER12));
	pArray->xltype = xltypeMulti | xlbitDLLFree;
	pArray->val.array.columns = 1;
	pArray->val.array.rows = NumTimeSteps;
	pArray->val.array.lparray = solution;
	return pArray;
}

LPXLOPER12 FullyImplicitXL(
	const double delta_time, const double TotSimTime, const double delta_depth, const double max_depth,
	const double effective_diffusion, const double effective_velocity, const double decay_rate,
	LPXLOPER12 &surface_tracer_concentrations, bool IsDescending) {
	

	const size_t NumTimeSteps = TotSimTime / delta_time;
	const size_t NumDepthSteps = max_depth / delta_depth + 1;
	// Construct factors in the iterative equation obtained from the finite
	// difference method. These parameters are described in the
	// solver_method.tex file.
	const double alpha = (effective_diffusion * delta_time) /
		(std::pow(delta_depth, 2));
	const double beta = (effective_velocity * delta_time) / (2 * delta_depth);
	// Construct the diagonal entries of the tridiagonal matrix.
	const double current_time_lower_diagonal = -alpha - beta;
	const double current_time_middle_diagonal = 1 + decay_rate + 2 * alpha;
	const double current_time_upper_diagonal = -alpha + beta;
	long i, j, start_i, end_i;
	LPXLOPER12 pT;
	// Initialize the previous time solution with boundary condition at t = 0
	std::vector<double> previous_time_solution(NumDepthSteps, 0.);
	if (IsDescending)
		pT = surface_tracer_concentrations->val.array.lparray + surface_tracer_concentrations->val.array.rows - 1;
	else
		pT = surface_tracer_concentrations->val.array.lparray; 
	previous_time_solution[0] = pT->val.num;
	// Find the depth step closest to the requested depth
	//size_t requested_depth_step = round(requested_depth / delta_depth);
	long XL_size = (NumTimeSteps)*(NumDepthSteps);
	XLOPER12 *solution = (XLOPER12 *)malloc(XL_size * sizeof(XLOPER12));
	//std::vector<double> solution(NumTimeSteps + 1);
	// Add initial value to solution;
	std::vector<double> previous_time_vector(NumDepthSteps-1);
	start_i = 0;
	for (size_t time_step = 1; time_step < NumTimeSteps+1; ++time_step) {
		// Calculate the RHS vector from the previous time step. This is simply the
		// previous time step solution, plus some boundary offsets.
		std::copy(std::next(previous_time_solution.begin(), 1),
			std::prev(previous_time_solution.end(), 1),
			previous_time_vector.begin());
		// Now add boundary condition offsets.
		// Note, the boundary offset at max depth is 0 because boundary at
		// max_depth is 0.
		if (time_step < surface_tracer_concentrations->val.array.rows) {
			if (IsDescending)
				pT = surface_tracer_concentrations->val.array.lparray + surface_tracer_concentrations->val.array.rows - time_step - 1;
			else
				pT = surface_tracer_concentrations->val.array.lparray + time_step;
		}
		previous_time_vector[0] -= (current_time_lower_diagonal *
			pT->val.num);

		// Calculate solution for this time step
		std::vector<double> current_time_solution =
			ThomasAlgorithimSingleValue(
				current_time_lower_diagonal, current_time_middle_diagonal,
				current_time_upper_diagonal, previous_time_vector);
		// Insert the solution into the previous_time_solution
		const auto insert_start = std::next(previous_time_solution.begin(), 1);
		std::copy(current_time_solution.begin(), current_time_solution.end(),
			insert_start);
		// Set surface boundary value of previous_time_solution
		previous_time_solution[0] = pT->val.num;
		// Add the relevant value to the wanted_depth_solution
		if (IsDescending) {
			for (i = 0; i < NumDepthSteps; i++)
			{
				solution[XL_size - start_i-1].xltype = xltypeNum;
				solution[XL_size - start_i-1].val.num = previous_time_solution[NumDepthSteps-i-1];
				start_i++;
			}
		}
		else {
			for (i = 0; i < NumDepthSteps; i++)
			{
				solution[start_i].xltype = xltypeNum;
				solution[start_i].val.num = previous_time_solution[i];
				start_i++;
			}
		}
	}
	// Create an array of pointers to XLOPER12 values.
	LPXLOPER12 pArray = (LPXLOPER12)malloc(XL_size * sizeof(LPXLOPER12));
	pArray->xltype = xltypeMulti | xlbitDLLFree;
	pArray->val.array.columns = NumDepthSteps;
	pArray->val.array.rows = NumTimeSteps;
	pArray->val.array.lparray = solution;
	return pArray;
}

//XLL Functions

HWND g_hWndMain = NULL;
HANDLE g_hInst = NULL;
HINSTANCE g_hInst2 = NULL;
XCHAR g_szBuffer[100] = L"";
#define MAX_V12_STRBUFFLEN    32678
#define g_rgWorksheetFuncsRows 5 //changed from 44 to match removal of C14 functions
#define g_rgWorksheetFuncsCols 29
static LPWSTR g_rgWorksheetFuncs
[g_rgWorksheetFuncsRows][g_rgWorksheetFuncsCols] =
{
	{ L"SolveNewtonMethod",
		L"QQQQIQQQQQKQQQQQQQBBBBBQIQQ",                   // up to 255 args in Excel 2007, 
			//QQQQIQQQQQKQQQQ							   // upto 29 args in Excel 2003 and earlier versions
		L"SolveNewtonMethod",
		L"MeasTracerConcs,MeasSigmas,SampleDates,ModelNum,FitParmIndexes,InitModVals,LoBounds,HiBounds,Tracers,DateRange,TracerInputRange,Lambdas,UZtime,UZtimeCond,TracerComp_2, DIC-1, DIC-2",
		L"1",
		L"TracerLPM Add-In",
		L"",                                    
		L"",                                  
		L"Solves for age and error",   
		L"Measured Tracer Concentrations",
		L"Measured Sigmas",
		L"Sample date(s)",
		L"Model to be used: DM(1),PFM(2),PEM(3),EPM(4),EMM(5),FDM(6),GAM(7)",
		L"Parameters to optimize: UZtime(1), Mean Age(2), 1st Model Parm(3), 2nd Model Parm(4), 1st Mixing Fraction(5), 1st Model Parm. 2nd deriv.(7), 2nd Model Parm. 2nd deriv.(8)",
		L"Initial Model Values",
		L"Low boundary condition",
		L"High boundary condition",
		L"Tracers to be included",
		L"Date range for samples (must be same number of dates as tracers)",
		L"Tracer range of samples (must be same number tracer values as dates)",
		L"Lambda values",
		L"Unsaturated zone time",
		L"Unsaturated zone time condition",
		L"Tracer component 2",
		L"Dissolved inorganic carbon content - component 1",
		L"Dissolved inorganic carbon content - component 2"
	},
	{ L"SolveLevenbergMarquardt",
		L"QQQQIQQQQQKQQQQQQQBBBBBQIQQ",                   // up to 255 args in Excel 2007, 
														  //QQQQIQQQQQKQQQQ							   // upto 29 args in Excel 2003 and earlier versions
		L"SolveLevenbergMarquardt",
		L"MeasTracerConcs,MeasSigmas,SampleDates,ModelNum,FitParmIndexes,InitModVals,LoBounds,HiBounds,Tracers,DateRange,TracerInputRange,Lambdas,UZtime,UZtimeCond,TracerComp_2, DIC-1, DIC-2",
		L"1",
		L"TracerLPM Add-In",
		L"",
		L"",
		L"Solves for age and error",
		L"Measured Tracer Concentrations",
		L"Measured Sigmas",
		L"Sample date(s)",
		L"Model to be used: DM(1),PFM(2),PEM(3),EPM(4),EMM(5),FDM(6),GAM(7)",
		L"Parameters to optimize: UZtime(1), Mean Age(2), 1st Model Parm(3), 2nd Model Parm(4), 1st Mixing Fraction(5), 1st Model Parm. 2nd deriv.(7), 2nd Model Parm. 2nd deriv.(8)",
		L"Initial Model Values",
		L"Low boundary condition",
		L"High boundary condition",
		L"Tracers to be included",
		L"Date range for samples (must be same number of dates as tracers)",
		L"Tracer range of samples (must be same number tracer values as dates)",
		L"Lambda values",
		L"Unsaturated zone time",
		L"Unsaturated zone time condition",
		L"Tracer component 2",
		L"Dissolved inorganic carbon content - component 1",
		L"Dissolved inorganic carbon content - component 2"
	},
	{ L"SolveGNLM",
	L"QQQQIQQQQQKQQQQQQQBBBBBQIQQ",                   // up to 255 args in Excel 2007, 
													  //QQQQIQQQQQKQQQQ							   // upto 29 args in Excel 2003 and earlier versions
	L"SolveGNLM",
	L"MeasTracerConcs,MeasSigmas,SampleDates,ModelNum,FitParmIndexes,InitModVals,LoBounds,HiBounds,Tracers,DateRange,TracerInputRange,Lambdas,UZtime,UZtimeCond,TracerComp_2, DIC-1, DIC-2",
	L"1",
	L"TracerLPM Add-In",
	L"",
	L"",
	L"Solves for age and error",
	L"Measured Tracer Concentrations",
	L"Measured Sigmas",
	L"Sample date(s)",
	L"Model to be used: DM(1),PFM(2),PEM(3),EPM(4),EMM(5),FDM(6),GAM(7)",
	L"Parameters to optimize: UZtime(1), Mean Age(2), 1st Model Parm(3), 2nd Model Parm(4), 1st Mixing Fraction(5), 1st Model Parm. 2nd deriv.(7), 2nd Model Parm. 2nd deriv.(8)",
	L"Initial Model Values",
	L"Low boundary condition",
	L"High boundary condition",
	L"Tracers to be included",
	L"Date range for samples (must be same number of dates as tracers)",
	L"Tracer range of samples (must be same number tracer values as dates)",
	L"Lambda values",
	L"Unsaturated zone time",
	L"Unsaturated zone time condition",
	L"Tracer component 2",
	L"Dissolved inorganic carbon content - component 1",
	L"Dissolved inorganic carbon content - component 2"
	},
	{ L"UZ_Solver1D",
		L"QEEEEEEEEQQ",                   // up to 255 args in Excel 2007, 
		L"UZ_Solver1D",
		L"delta_time,TotSimTime,delta_depth,max_depth,effective_diffusion,effective_velocity,decay_rate,requested_depth,tracer_date_range,surface_tracer_concs",
		L"1",
		L"TracerLPM Add-In",
		L"",
		L"",
		L"Calculates the time-series concentrations of a gas tracer for a single water table depth",
		L"Delta Time (double)",
		L"Total simulation time (double)",
		L"Delta depth (double)",
		L"Max Depth (double)",
		L"Effective Diffusion (double)",
		L"Effective Velocity (double)",
		L"Decay Rate (double)",
		L"Requested Depth (double)",
		L"Number of concentration values (int)",
		L"Date range of tracers concentrations (Range)",
		L"Surface Tracer Concentrations (Range)",
	},
	{ L"UZ_Solver2D",
	L"QEEEEEEEQQ",                   // up to 255 args in Excel 2007, 
		L"UZ_Solver2D",
		L"delta_time,TotSimTime,delta_depth,max_depth,effective_diffusion,effective_velocity,decay_rate,tracer_date_range,surface_tracer_concs",
		L"1",
		L"TracerLPM Add-In",
		L"",
		L"",
		L"Calculates the time-series concentrations of a gas tracer for different water table depths",
		L"Delta Time (double)",
		L"Total simulation time (double)",
		L"Delta depth (double)",
		L"Max Depth (double)",
		L"Effective Diffusion (double)",
		L"Effective Velocity (double)",
		L"Decay Rate (double)",
		L"Number of concentration values (int)",
		L"Date range of tracers concentrations (Range)",
		L"Surface Tracer Concentrations (Range)",
	}
};

//**************************************************************************

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

wchar_t * deep_copy_wcs(const wchar_t *p_source)
{
    if(!p_source)
        return NULL;

    size_t source_len = p_source[0];
    bool truncated = false;

    if(source_len >= MAX_V12_STRBUFFLEN)
    {
        source_len = MAX_V12_STRBUFFLEN - 1; // Truncate the copy
        truncated = true;
    }
    wchar_t *p_copy = new wchar_t[source_len + 1];
    wcsncpy_s(p_copy, source_len+1, p_source+1, source_len);
	//wcsncpy(p_copy, p_source+1, source_len + 1);
    if(truncated)
        p_copy[0] = source_len;
    return p_copy;
}
//***************************************************************************
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
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][24]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][25]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][26]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][27]),
			  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][28]));
	}
	
	// Free the XLL filename //
	Excel12f(xlFree, 0, 2, (LPXLOPER12) &xTest, (LPXLOPER12) &xDLL);

	return 1;
}



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
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][23]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][24]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][25]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][26]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][27]),
				  (LPXLOPER12) TempStr12(g_rgWorksheetFuncs[i][28]));
			/// Free oper returned by xl //
			Excel12f(xlFree, 0, 1, (LPXLOPER12) &xDLL);

			return(LPXLOPER12) &xRegId;
		}
	}
	return 0;
}

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

__declspec(dllexport) int WINAPI xlAutoRemove(void)
{
	// Show a dialog box indicating that the XLL was successfully removed //
	//Excel12f(xlcAlert, 0, 2, TempStr12(L"Thank you for removing TracerLPMfunctions.XLL!"),
	//	  TempInt12(2));
	return 1;
}

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
		xInfo.val.str = L"\016TracerLPM_2015";
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

void WINAPI xlAutoFree12(LPXLOPER12 pxFree)
{
    if(pxFree->xltype & xltypeMulti)
    {
// Assume all string elements were allocated using malloc, and
// need to be freed using free. Then free the array itself.
        int size = pxFree->val.array.rows *
            pxFree->val.array.columns;
        LPXLOPER12 p = pxFree->val.array.lparray;

        for(; size-- > 0; p++) // check elements for strings
            if(p->xltype == xltypeStr)
                free(p->val.str);

        free(pxFree->val.array.lparray);
    }
    else if(pxFree->xltype & xltypeStr)
    {
        free(pxFree->val.str);
    }
    else if(pxFree->xltype & xltypeRef)
    {
        free(pxFree->val.mref.lpmref);
    }
// Assume pxFree was itself dynamically allocated using malloc.
    free(pxFree);
}

#endif