

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Mon Jan 02 12:00:34 2017
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

#ifndef __IInterpolateValue_FWD_DEFINED__
#define __IInterpolateValue_FWD_DEFINED__
typedef interface IInterpolateValue IInterpolateValue;

#endif 	/* __IInterpolateValue_FWD_DEFINED__ */


#ifndef ___InterpolateValue_FWD_DEFINED__
#define ___InterpolateValue_FWD_DEFINED__
typedef interface _InterpolateValue _InterpolateValue;

#endif 	/* ___InterpolateValue_FWD_DEFINED__ */


#ifndef __InterpolateValue_FWD_DEFINED__
#define __InterpolateValue_FWD_DEFINED__

#ifdef __cplusplus
typedef class InterpolateValue InterpolateValue;
#else
typedef struct InterpolateValue InterpolateValue;
#endif /* __cplusplus */

#endif 	/* __InterpolateValue_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 


/* interface __MIDL_itf_MyIDL_0000_0000 */
/* [local] */ 

#pragma once


extern RPC_IF_HANDLE __MIDL_itf_MyIDL_0000_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_MyIDL_0000_0000_v0_0_s_ifspec;


#ifndef __InterpolateValue_LIBRARY_DEFINED__
#define __InterpolateValue_LIBRARY_DEFINED__

/* library InterpolateValue */
/* [version][helpstring][uuid] */ 


EXTERN_C const IID LIBID_InterpolateValue;

#ifndef __IInterpolateValue_DISPINTERFACE_DEFINED__
#define __IInterpolateValue_DISPINTERFACE_DEFINED__

/* dispinterface IInterpolateValue */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID_IInterpolateValue;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("C1A9859A-1480-44BC-885A-667F8F0E0164")
    IInterpolateValue : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct IInterpolateValueVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IInterpolateValue * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IInterpolateValue * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IInterpolateValue * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IInterpolateValue * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IInterpolateValue * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IInterpolateValue * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IInterpolateValue * This,
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
    } IInterpolateValueVtbl;

    interface IInterpolateValue
    {
        CONST_VTBL struct IInterpolateValueVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IInterpolateValue_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IInterpolateValue_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IInterpolateValue_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IInterpolateValue_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IInterpolateValue_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IInterpolateValue_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IInterpolateValue_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* __IInterpolateValue_DISPINTERFACE_DEFINED__ */


#ifndef ___InterpolateValue_DISPINTERFACE_DEFINED__
#define ___InterpolateValue_DISPINTERFACE_DEFINED__

/* dispinterface _InterpolateValue */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__InterpolateValue;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("D13EBD58-ADFE-4F31-89B2-FBBEB569B4B5")
    _InterpolateValue : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _InterpolateValueVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _InterpolateValue * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _InterpolateValue * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _InterpolateValue * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _InterpolateValue * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _InterpolateValue * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _InterpolateValue * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _InterpolateValue * This,
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
    } _InterpolateValueVtbl;

    interface _InterpolateValue
    {
        CONST_VTBL struct _InterpolateValueVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _InterpolateValue_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _InterpolateValue_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _InterpolateValue_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _InterpolateValue_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _InterpolateValue_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _InterpolateValue_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _InterpolateValue_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___InterpolateValue_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_InterpolateValue;

#ifdef __cplusplus

class DECLSPEC_UUID("86201DD6-BAE5-493C-AB74-4CE7F1344BEB")
InterpolateValue;
#endif
#endif /* __InterpolateValue_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


