#include "stdafx.h"


// interpolate a value

// Numerical recipes routines
#include "nrutil.h"

void polint(float xa[], float ya[], int n, float x, float *y, float *dy);
void locate(float xx[], unsigned long n, float x, unsigned long *j);

// interpolate
float interpolateValue(
	float		*	px,
	float		*	py,
	unsigned long	n,
	double			xValue)
{
	float x = (float)xValue;
	float y = (float) 0.0;					// returned value
	float dy;
	unsigned long j;
	unsigned long m = 4;						// 4 point polynomial fit
	unsigned long k;
	// use locate to find index of interest
	locate(px, n, x, &j);
	k = IMIN(IMAX((j - m - 1) / 2, 1), n + 1 - m);
	polint(&px[k - 1], &py[k - 1], m, x, &y, &dy);
	return y;
}
