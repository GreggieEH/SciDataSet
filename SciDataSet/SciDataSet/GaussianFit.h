#pragma once
class CGaussianFit
{
public:
	CGaussianFit();
	~CGaussianFit();
	void			SetPeakThreshold(
						double		PeakThreshold);
	BOOL			SetData(
						long		nArray,
						double	*	pX,
						double	*	pY);
	BOOL			FitPeaks();
	long			GetNumPeaks();
	double			GetPeakPosition(
						long		index);
	double			GetPeakHeight(
						long		index);
	double			GetPeakWidth(
						long		index);
	void			DisplayPeaks(
						HWND		hwndParent);
private:
	IDispatch	*	m_pdisp;
};

