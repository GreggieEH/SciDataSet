

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Tue Oct 04 14:34:48 2016
 */
/* Compiler settings for MyIDL.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.00.0603 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__


#ifndef __MyIDL_h_h__
#define __MyIDL_h_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __ISciDataSet_FWD_DEFINED__
#define __ISciDataSet_FWD_DEFINED__
typedef interface ISciDataSet ISciDataSet;

#endif 	/* __ISciDataSet_FWD_DEFINED__ */


#ifndef ___SciDataSet_FWD_DEFINED__
#define ___SciDataSet_FWD_DEFINED__
typedef interface _SciDataSet _SciDataSet;

#endif 	/* ___SciDataSet_FWD_DEFINED__ */


#ifndef __SciDataSet_FWD_DEFINED__
#define __SciDataSet_FWD_DEFINED__

#ifdef __cplusplus
typedef class SciDataSet SciDataSet;
#else
typedef struct SciDataSet SciDataSet;
#endif /* __cplusplus */

#endif 	/* __SciDataSet_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 


/* interface __MIDL_itf_MyIDL_0000_0000 */
/* [local] */ 

#pragma once


extern RPC_IF_HANDLE __MIDL_itf_MyIDL_0000_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_MyIDL_0000_0000_v0_0_s_ifspec;


#ifndef __SciDataSet_LIBRARY_DEFINED__
#define __SciDataSet_LIBRARY_DEFINED__

/* library SciDataSet */
/* [version][helpstring][uuid] */ 


EXTERN_C const IID LIBID_SciDataSet;

#ifndef __ISciDataSet_DISPINTERFACE_DEFINED__
#define __ISciDataSet_DISPINTERFACE_DEFINED__

/* dispinterface ISciDataSet */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID_ISciDataSet;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("E2DEA9B5-F6F0-440d-B728-D5A5EA15CB42")
    ISciDataSet : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct ISciDataSetVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ISciDataSet * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ISciDataSet * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ISciDataSet * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ISciDataSet * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ISciDataSet * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ISciDataSet * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ISciDataSet * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } ISciDataSetVtbl;

    interface ISciDataSet
    {
        CONST_VTBL struct ISciDataSetVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ISciDataSet_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ISciDataSet_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ISciDataSet_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ISciDataSet_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ISciDataSet_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ISciDataSet_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ISciDataSet_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __ISciDataSet_DISPINTERFACE_DEFINED__ */


#ifndef ___SciDataSet_DISPINTERFACE_DEFINED__
#define ___SciDataSet_DISPINTERFACE_DEFINED__

/* dispinterface _SciDataSet */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__SciDataSet;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("95A4AD42-7631-436a-9B07-274C3D2D8470")
    _SciDataSet : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _SciDataSetVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _SciDataSet * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _SciDataSet * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _SciDataSet * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _SciDataSet * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _SciDataSet * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _SciDataSet * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _SciDataSet * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } _SciDataSetVtbl;

    interface _SciDataSet
    {
        CONST_VTBL struct _SciDataSetVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _SciDataSet_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _SciDataSet_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _SciDataSet_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _SciDataSet_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _SciDataSet_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _SciDataSet_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _SciDataSet_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___SciDataSet_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_SciDataSet;

#ifdef __cplusplus

class DECLSPEC_UUID("53C6AA39-8826-4572-A514-05105800D22F")
SciDataSet;
#endif
#endif /* __SciDataSet_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


