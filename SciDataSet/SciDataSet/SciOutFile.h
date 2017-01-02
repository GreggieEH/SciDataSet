#pragma once

class CSciOutFile
{
public:
	CSciOutFile(void);
	~CSciOutFile(void);
	BOOL			GetMyObject(
						IDispatch	**	ppdisp);
	void			SetfileName(
						LPCTSTR			szfileName);
	void			SettimeStamp(
						VARIANT		*	timeStamp);
	void			SetsourceFile(
						LPCTSTR			szsourceFile);
	void			SetcalibrationFile(
						LPCTSTR			szcalibrationFile);
	void			SetanalysisType(
						LPCTSTR			szanalysisType);
	void			SetxValues(
						long			nValues,
						double		*	pxValues);
	void			SetyValues(
						long			nValues,
						double		*	pyValues);
	void			saveFile();
	void			SetComment(
						LPCTSTR			szComment);
	void			SetyUnits(
						LPCTSTR			szyUnits);
	void			SetExtraData(
						LPCTSTR			szTitle,
						long			nArray,
						double		*	pWaves,
						double		*	pValues);
protected:
	BOOL			Get_clsIDataSet(
						IDispatch	**	ppdisp);
private:
	IDispatch	*	m_pdisp;
};
