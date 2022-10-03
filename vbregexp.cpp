/*
 * Modified from https://github.com/Gcenx/winecx's vbregexp.c
 *
 * COM server registration follows the example given in John's Blog:
 * https://nachtimwald.com/2012/04/08/wrapping-a-c-library-in-comactivex/
 *
 * Copyright 2013 Piotr Caban for CodeWeavers
 *
 * SPDX-License-Identifier: LGPL-2.1-or-later
 */

#include <windows.h>
#include <vld.h>
#define RE_PREFIX srell
#include "srell.hpp"
#include "vbsregexp55_h.h"
#include "vbsregexp55_i.c"

enum {
    REG_FOLD = RE_PREFIX::wregex::icase,
    REG_MULTILINE = RE_PREFIX::wregex::multiline,
    REG_GLOB = 0x4000,
    REG_PATTERN = 0x8000,
};

// Provide some dummies for wine's debugging facilities

#define TRACE sizeof
#define ERR sizeof
#define FIXME sizeof

static char const *debugstr_guid(const GUID &);
static char const *wine_dbgstr_w(WCHAR const *);
static char const *debugstr_w(WCHAR const *);
static char const *debugstr_variant(VARIANT const *);

static inline BOOL is_digit(WCHAR c)
{
    return '0' <= c && c <= '9';
}

#define REGEXP_TID_LIST \
    XDIID(RegExp2), \
    XDIID(Match2), \
    XDIID(MatchCollection2), \
    XDIID(SubMatches)

typedef enum {
#define XDIID(iface) iface ## _tid
    REGEXP_TID_LIST,
#undef XDIID
    REGEXP_LAST_tid
} regexp_tid_t;

static IID tid_ids[] = {
#define XDIID(iface) IID_I ## iface
    REGEXP_TID_LIST
#undef XDIID
};

HMODULE g_module = NULL;

static ITypeLib *typelib;
static ITypeInfo *typeinfos[REGEXP_LAST_tid];

static HRESULT init_regexp_typeinfo(regexp_tid_t tid)
{
    HRESULT hres;

    if (!typelib) {
        ITypeLib *tl;

        WCHAR szFileName[MAX_PATH];
        GetModuleFileNameW(g_module, szFileName, MAX_PATH);
        hres = LoadTypeLib(szFileName, &tl);
        if (FAILED(hres)) {
            ERR("LoadRegTypeLib failed: %08lx\n", hres);
            return hres;
        }

        if (InterlockedCompareExchangePointer(reinterpret_cast<void **>(&typelib), tl, NULL))
            tl->Release();
    }

    if (!typeinfos[tid]) {
        ITypeInfo *ti;

        hres = typelib->GetTypeInfoOfGuid(tid_ids[tid], &ti);
        if (FAILED(hres)) {
            ERR("GetTypeInfoOfGuid(%s) failed: %08lx\n", debugstr_guid(tid_ids[tid]), hres);
            return hres;
        }

        if (InterlockedCompareExchangePointer(reinterpret_cast<void **>(typeinfos + tid), ti, NULL))
            ti->Release();
    }

    return S_OK;
}

template<class Self>
class ZeroInit
{
protected:
    ZeroInit() {
        memset(static_cast<Self *>(this), 0, sizeof(Self));
    }
};

class SubMatchesEnum
    : public ZeroInit<SubMatchesEnum>
    , public IEnumVARIANT
{
private:
    LONG ref;
    ISubMatches *sm;
    LONG pos;
    LONG count;
public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv);
    ULONG STDMETHODCALLTYPE AddRef();
    ULONG STDMETHODCALLTYPE Release();

    HRESULT STDMETHODCALLTYPE Next(ULONG celt, VARIANT *rgVar, ULONG *pCeltFetched);
    HRESULT STDMETHODCALLTYPE Skip(ULONG celt);
    HRESULT STDMETHODCALLTYPE Reset();
    HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT **ppEnum);

    static HRESULT create(ISubMatches *sm, IUnknown **ppEnum);
};

class SubMatches
    : public ZeroInit<SubMatches>
    , public ISubMatches
{
private:
    LONG ref;
public:
    RE_PREFIX::wcmatch result;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv);
    ULONG STDMETHODCALLTYPE AddRef();
    ULONG STDMETHODCALLTYPE Release();

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo);
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(
        REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember,
                                     REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams,
                                     VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

    HRESULT STDMETHODCALLTYPE get_Item(LONG index, VARIANT *pSubMatch);
    HRESULT STDMETHODCALLTYPE get_Count(LONG *pCount);
    HRESULT STDMETHODCALLTYPE get__NewEnum(IUnknown **ppEnum);

    static HRESULT create(RE_PREFIX::wcmatch &result, SubMatches **sub_matches);
};

class Match2
    : public ZeroInit<Match2>
    , public IMatch2
    , public IMatch
{
private:
    LONG ref;
public:
    DWORD index;
    std::wstring value;
    SubMatches *sub_matches;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv);
    ULONG STDMETHODCALLTYPE AddRef();
    ULONG STDMETHODCALLTYPE Release();

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo);
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(
        REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember,
                                     REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams,
                                     VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

    HRESULT STDMETHODCALLTYPE get_Value(BSTR *pValue);
    HRESULT STDMETHODCALLTYPE get_FirstIndex(LONG *pFirstIndex);
    HRESULT STDMETHODCALLTYPE get_Length(LONG *pLength);
    HRESULT STDMETHODCALLTYPE get_SubMatches(IDispatch **ppSubMatches);

    static HRESULT create(DWORD pos, RE_PREFIX::wcmatch &result, IMatch2 **match);
};

class MatchCollectionEnum
    : public ZeroInit<MatchCollectionEnum>
    , public IEnumVARIANT
{
private:
    LONG ref;
    IMatchCollection2 *mc;
    LONG pos;
    LONG count;
public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv);
    ULONG STDMETHODCALLTYPE AddRef();
    ULONG STDMETHODCALLTYPE Release();

    HRESULT STDMETHODCALLTYPE Next(ULONG celt, VARIANT *rgVar, ULONG *pCeltFetched);
    HRESULT STDMETHODCALLTYPE Skip(ULONG celt);
    HRESULT STDMETHODCALLTYPE Reset();
    HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT **ppEnum);

    static HRESULT create(IMatchCollection2 *mc, IUnknown **ppEnum);
};

class MatchCollection2
    : public ZeroInit<MatchCollection2>
    , public IMatchCollection2
    , public IMatchCollection
{
private:
    LONG ref;
public:
    std::vector<IMatch2 *> matches;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv);
    ULONG STDMETHODCALLTYPE AddRef();
    ULONG STDMETHODCALLTYPE Release();

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo);
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(
        REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember,
                                     REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams,
                                     VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

    HRESULT STDMETHODCALLTYPE get_Item(LONG index, IDispatch **ppMatch);
    HRESULT STDMETHODCALLTYPE get_Count(LONG *pCount);
    HRESULT STDMETHODCALLTYPE get__NewEnum(IUnknown **ppEnum);

    static HRESULT create(MatchCollection2 **match_collection);
};

class RegExp2
    : public ZeroInit<RegExp2>
    , public IRegExp2
    , public IRegExp
{
private:
    LONG ref;
public:
    std::wstring pattern;
    RE_PREFIX::wregex regexp;

    WORD flags;
    WORD state;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv);
    ULONG STDMETHODCALLTYPE AddRef();
    ULONG STDMETHODCALLTYPE Release();

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo);
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(
        REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember,
                                     REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams,
                                     VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

    HRESULT STDMETHODCALLTYPE get_Pattern(BSTR *);
    HRESULT STDMETHODCALLTYPE put_Pattern(BSTR);
    HRESULT STDMETHODCALLTYPE get_IgnoreCase(VARIANT_BOOL *);
    HRESULT STDMETHODCALLTYPE put_IgnoreCase(VARIANT_BOOL);
    HRESULT STDMETHODCALLTYPE get_Global(VARIANT_BOOL *);
    HRESULT STDMETHODCALLTYPE put_Global(VARIANT_BOOL);
    HRESULT STDMETHODCALLTYPE get_Multiline(VARIANT_BOOL *);
    HRESULT STDMETHODCALLTYPE put_Multiline(VARIANT_BOOL);
    HRESULT STDMETHODCALLTYPE Execute(BSTR source, IDispatch **ppMatches);
    HRESULT STDMETHODCALLTYPE Test(BSTR source, VARIANT_BOOL *pMatch);
    HRESULT STDMETHODCALLTYPE Replace(BSTR source, VARIANT replaceVar, BSTR *pDestString);
    HRESULT STDMETHODCALLTYPE Replace(BSTR source, BSTR replace, BSTR *pDestString);

    static HRESULT create(IDispatch **ret);

private:
    void update();
};

static class RegExp2Factory : public IClassFactory
{
public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID, void **);
    ULONG STDMETHODCALLTYPE AddRef();
    ULONG STDMETHODCALLTYPE Release();

    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown *, REFIID, void **);
    HRESULT STDMETHODCALLTYPE LockServer(BOOL);
} gRegExp2Factory;

HRESULT STDMETHODCALLTYPE SubMatchesEnum::QueryInterface(REFIID riid, void **ppv)
{
    if (IsEqualGUID(riid, IID_IUnknown)) {
        TRACE("(%p)->(IID_IUnknown %p)\n", this, ppv);
        *ppv = static_cast<IEnumVARIANT *>(this);
    } else if (IsEqualGUID(riid, IID_IEnumVARIANT)) {
        TRACE("(%p)->(IID_IEnumVARIANT %p)\n", this, ppv);
        *ppv = static_cast<IEnumVARIANT *>(this);
    } else {
        FIXME("(%p)->(%s %p)\n", this, debugstr_guid(riid), ppv);
        *ppv = NULL;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

ULONG STDMETHODCALLTYPE SubMatchesEnum::AddRef()
{
    LONG const ref = InterlockedIncrement(&this->ref);
    TRACE("(%p) ref=%ld\n", this, ref);
    return ref;
}

ULONG STDMETHODCALLTYPE SubMatchesEnum::Release()
{
    LONG const ref = InterlockedDecrement(&this->ref);
    TRACE("(%p) ref=%ld\n", this, ref);
    if (!ref) {
        sm->Release();
        delete this;
    }
    return ref;
}

HRESULT STDMETHODCALLTYPE SubMatchesEnum::Next(ULONG celt, VARIANT *rgVar, ULONG *pCeltFetched)
{
    LONG i;
    HRESULT hres = S_OK;

    TRACE("(%p)->(%lu %p %p)\n", this, celt, rgVar, pCeltFetched);

    if (pos >= count) {
        if (pCeltFetched)
            *pCeltFetched = 0;
        return S_FALSE;
    }

    for (i = 0; i < static_cast<LONG>(celt) && pos + i < count; i++) {
        hres = sm->get_Item(pos + i, rgVar + i);
        if (FAILED(hres))
            break;
    }
    if (FAILED(hres)) {
        while (i--)
            VariantClear(rgVar + i);
        return hres;
    }

    if (pCeltFetched)
        *pCeltFetched = i;
    pos += i;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE SubMatchesEnum::Skip(ULONG celt)
{
    TRACE("(%p)->(%lu)\n", this, celt);

    if (pos + static_cast<LONG>(celt) <= count)
        pos += celt;
    else
        pos = count;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE SubMatchesEnum::Reset()
{
    TRACE("(%p)\n", this);

    pos = 0;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE SubMatchesEnum::Clone(IEnumVARIANT **ppEnum)
{
    FIXME("(%p)->(%p)\n", this, ppEnum);
    return E_NOTIMPL;
}

HRESULT SubMatchesEnum::create(ISubMatches *sm, IUnknown **ppEnum)
{
    SubMatchesEnum *ret = new(std::nothrow) SubMatchesEnum;

    if (!ret)
        return E_OUTOFMEMORY;

    ret->ref = 1;
    sm->get_Count(&ret->count);
    ret->sm = sm;
    sm->AddRef();

    *ppEnum = ret;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE SubMatches::QueryInterface(REFIID riid, void **ppv)
{
    if (IsEqualGUID(riid, IID_IUnknown)) {
        TRACE("(%p)->(IID_IUnknown %p)\n", this, ppv);
        *ppv = static_cast<ISubMatches *>(this);
    } else if (IsEqualGUID(riid, IID_IDispatch)) {
        TRACE("(%p)->(IID_IDispatch %p)\n", this, ppv);
        *ppv = static_cast<ISubMatches *>(this);
    } else if (IsEqualGUID(riid, IID_ISubMatches)) {
        TRACE("(%p)->(IID_ISubMatches %p)\n", this, ppv);
        *ppv = static_cast<ISubMatches *>(this);
    } else {
        FIXME("(%p)->(%s %p)\n", this, debugstr_guid(riid), ppv);
        *ppv = NULL;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

ULONG STDMETHODCALLTYPE SubMatches::AddRef()
{
    LONG const ref = InterlockedIncrement(&this->ref);
    TRACE("(%p) ref=%ld\n", this, ref);
    return ref;
}

ULONG STDMETHODCALLTYPE SubMatches::Release()
{
    LONG const ref = InterlockedDecrement(&this->ref);
    TRACE("(%p) ref=%ld\n", this, ref);
    if (!ref) {
        delete this;
    }
    return ref;
}

HRESULT STDMETHODCALLTYPE SubMatches::GetTypeInfoCount(UINT *pctinfo)
{
    TRACE("(%p)->(%p)\n", this, pctinfo);
    *pctinfo = 1;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE SubMatches::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
    TRACE("(%p)->(%u %lu %p)\n", this, iTInfo, lcid, ppTInfo);
    if (iTInfo != 0)
        return DISP_E_BADINDEX;
    (*ppTInfo = typeinfos[SubMatches_tid])->AddRef();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE SubMatches::GetIDsOfNames(
    REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
    TRACE("(%p)->(%s %p %u %lu %p)\n", this, debugstr_guid(riid),
          rgszNames, cNames, lcid, rgDispId);

    return typeinfos[SubMatches_tid]->GetIDsOfNames(rgszNames, cNames, rgDispId);
}

HRESULT STDMETHODCALLTYPE SubMatches::Invoke(DISPID dispIdMember,
        REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams,
        VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
    TRACE("(%p)->(%ld %s %ld %d %p %p %p %p)\n", this, dispIdMember, debugstr_guid(riid),
          lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);

    return typeinfos[SubMatches_tid]->Invoke(this, dispIdMember, wFlags,
            pDispParams, pVarResult, pExcepInfo, puArgErr);
}

HRESULT STDMETHODCALLTYPE SubMatches::get_Item(LONG index, VARIANT *pSubMatch)
{
    TRACE("(%p)->(%ld %p)\n", this, index, pSubMatch);

    if (!pSubMatch)
        return E_POINTER;

    if (index < 0 || index >= static_cast<LONG>(result.size() - 1))
        return E_INVALIDARG;

    RE_PREFIX::wcmatch::const_reference sm = result[index + 1];
    V_VT(pSubMatch) = VT_BSTR;
    V_BSTR(pSubMatch) = SysAllocStringLen(sm.first, sm.second - sm.first);

    if (!V_BSTR(pSubMatch))
        return E_OUTOFMEMORY;

    return S_OK;
}

HRESULT STDMETHODCALLTYPE SubMatches::get_Count(LONG *pCount)
{
    TRACE("(%p)->(%p)\n", this, pCount);

    if (!pCount)
        return E_POINTER;

    *pCount = result.size() - 1;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE SubMatches::get__NewEnum(IUnknown **ppEnum)
{
    FIXME("(%p)->(%p)\n", this, ppEnum);

    if (!ppEnum)
        return E_POINTER;

    return SubMatchesEnum::create(this, ppEnum);
}

HRESULT SubMatches::create(RE_PREFIX::wcmatch &result, SubMatches **sub_matches)
{
    HRESULT hres = init_regexp_typeinfo(SubMatches_tid);
    if (FAILED(hres))
        return hres;

    SubMatches *ret = new(std::nothrow) SubMatches;
    if (!ret)
        return E_OUTOFMEMORY;

    ret->result = result;

    ret->ref = 1;
    *sub_matches = ret;
    return hres;
}

HRESULT STDMETHODCALLTYPE Match2::QueryInterface(REFIID riid, void **ppv)
{
    if (IsEqualGUID(riid, IID_IUnknown)) {
        TRACE("(%p)->(IID_IUnknown %p)\n", this, ppv);
        *ppv = static_cast<IMatch2 *>(this);
    } else if (IsEqualGUID(riid, IID_IDispatch)) {
        TRACE("(%p)->(IID_IDispatch %p)\n", this, ppv);
        *ppv = static_cast<IMatch2 *>(this);
    } else if (IsEqualGUID(riid, IID_IMatch2)) {
        TRACE("(%p)->(IID_IMatch2 %p)\n", this, ppv);
        *ppv = static_cast<IMatch2 *>(this);
    } else if (IsEqualGUID(riid, IID_IMatch)) {
        TRACE("(%p)->(IID_IMatch %p)\n", this, ppv);
        *ppv = static_cast<IMatch *>(this);
    } else {
        FIXME("(%p)->(%s %p)\n", this, debugstr_guid(riid), ppv);
        *ppv = NULL;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

ULONG STDMETHODCALLTYPE Match2::AddRef()
{
    LONG const ref = InterlockedIncrement(&this->ref);
    TRACE("(%p) ref=%ld\n", this, ref);
    return ref;
}

ULONG STDMETHODCALLTYPE Match2::Release()
{
    LONG const ref = InterlockedDecrement(&this->ref);
    TRACE("(%p) ref=%ld\n", this, ref);
    if (!ref) {
        sub_matches->Release();
        delete this;
    }
    return ref;
}

HRESULT STDMETHODCALLTYPE Match2::GetTypeInfoCount(UINT *pctinfo)
{
    TRACE("(%p)->(%p)\n", this, pctinfo);
    *pctinfo = 1;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE Match2::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
    TRACE("(%p)->(%u %lu %p)\n", this, iTInfo, lcid, ppTInfo);
    if (iTInfo != 0)
        return DISP_E_BADINDEX;
    (*ppTInfo = typeinfos[Match2_tid])->AddRef();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE Match2::GetIDsOfNames(
    REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
    TRACE("(%p)->(%s %p %u %lu %p)\n", this, debugstr_guid(riid),
          rgszNames, cNames, lcid, rgDispId);

    return typeinfos[Match2_tid]->GetIDsOfNames(rgszNames, cNames, rgDispId);
}

HRESULT STDMETHODCALLTYPE Match2::Invoke(DISPID dispIdMember,
        REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams,
        VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
    TRACE("(%p)->(%ld %s %ld %d %p %p %p %p)\n", this, dispIdMember, debugstr_guid(riid),
          lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);

    return typeinfos[Match2_tid]->Invoke(static_cast<IMatch2 *>(this), dispIdMember, wFlags,
                                         pDispParams, pVarResult, pExcepInfo, puArgErr);
}

HRESULT STDMETHODCALLTYPE Match2::get_Value(BSTR *pValue)
{
    TRACE("(%p)->(%p)\n", this, pValue);

    if (!pValue)
        return E_POINTER;

    *pValue = SysAllocStringLen(value.c_str(), value.length());
    return *pValue ? S_OK : E_OUTOFMEMORY;
}

HRESULT STDMETHODCALLTYPE Match2::get_FirstIndex(LONG *pFirstIndex)
{
    TRACE("(%p)->(%p)\n", this, pFirstIndex);

    if (!pFirstIndex)
        return E_POINTER;

    *pFirstIndex = index;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE Match2::get_Length(LONG *pLength)
{
    TRACE("(%p)->(%p)\n", this, pLength);

    if (!pLength)
        return E_POINTER;

    *pLength = sub_matches->result.size();

    return S_OK;
}

HRESULT STDMETHODCALLTYPE Match2::get_SubMatches(IDispatch **ppSubMatches)
{
    TRACE("(%p)->(%p)\n", this, ppSubMatches);

    if (!ppSubMatches)
        return E_POINTER;

    *ppSubMatches = sub_matches;
    sub_matches->AddRef();
    return S_OK;
}

HRESULT Match2::create(DWORD pos, RE_PREFIX::wcmatch &result, IMatch2 **match)
{
    HRESULT hres = init_regexp_typeinfo(Match2_tid);
    if (FAILED(hres))
        return hres;

    Match2 *ret = new(std::nothrow) Match2;
    if (!ret)
        return E_OUTOFMEMORY;

    RE_PREFIX::wcmatch::const_reference sm = result[0];
    ret->index = pos;
    ret->value.assign(sm.first, sm.second);

    hres = SubMatches::create(result, &ret->sub_matches);
    if (FAILED(hres)) {
        delete ret;
        return hres;
    }

    ret->ref = 1;
    *match = ret;
    return hres;
}

HRESULT STDMETHODCALLTYPE MatchCollectionEnum::QueryInterface(REFIID riid, void **ppv)
{
    if (IsEqualGUID(riid, IID_IUnknown)) {
        TRACE("(%p)->(IID_IUnknown %p)\n", this, ppv);
        *ppv = static_cast<IEnumVARIANT *>(this);
    } else if (IsEqualGUID(riid, IID_IEnumVARIANT)) {
        TRACE("(%p)->(IID_IEnumVARIANT %p)\n", this, ppv);
        *ppv = static_cast<IEnumVARIANT *>(this);
    } else {
        FIXME("(%p)->(%s %p)\n", this, debugstr_guid(riid), ppv);
        *ppv = NULL;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

ULONG STDMETHODCALLTYPE MatchCollectionEnum::AddRef()
{
    LONG const ref = InterlockedIncrement(&this->ref);
    TRACE("(%p) ref=%ld\n", this, ref);
    return ref;
}

ULONG STDMETHODCALLTYPE MatchCollectionEnum::Release()
{
    LONG const ref = InterlockedDecrement(&this->ref);
    TRACE("(%p) ref=%ld\n", this, ref);
    if (!ref) {
        mc->Release();
        delete this;
    }
    return ref;
}

HRESULT STDMETHODCALLTYPE MatchCollectionEnum::Next(ULONG celt, VARIANT *rgVar, ULONG *pCeltFetched)
{
    LONG i;
    HRESULT hres = S_OK;

    TRACE("(%p)->(%lu %p %p)\n", this, celt, rgVar, pCeltFetched);

    if (pos >= count) {
        if (pCeltFetched)
            *pCeltFetched = 0;
        return S_FALSE;
    }

    for (i = 0; i < static_cast<LONG>(celt) && pos + i < count; i++) {
        V_VT(rgVar+i) = VT_DISPATCH;
        hres = mc->get_Item(pos + i, &V_DISPATCH(rgVar + i));
        if (FAILED(hres))
            break;
    }
    if (FAILED(hres)) {
        while (i--)
            VariantClear(rgVar + i);
        return hres;
    }

    if (pCeltFetched)
        *pCeltFetched = i;
    pos += i;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE MatchCollectionEnum::Skip(ULONG celt)
{
    TRACE("(%p)->(%lu)\n", this, celt);

    if (pos + static_cast<LONG>(celt) <= count)
        pos += celt;
    else
        pos = count;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE MatchCollectionEnum::Reset()
{
    TRACE("(%p)\n", this);

    pos = 0;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE MatchCollectionEnum::Clone(IEnumVARIANT **ppEnum)
{
    FIXME("(%p)->(%p)\n", this, ppEnum);
    return E_NOTIMPL;
}

HRESULT MatchCollectionEnum::create(IMatchCollection2 *mc, IUnknown **ppEnum)
{
    MatchCollectionEnum *ret = new(std::nothrow) MatchCollectionEnum;

    if (!ret)
        return E_OUTOFMEMORY;

    ret->ref = 1;
    mc->get_Count(&ret->count);
    ret->mc = mc;
    mc->AddRef();

    *ppEnum = ret;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE MatchCollection2::QueryInterface(REFIID riid, void **ppv)
{
    if (IsEqualGUID(riid, IID_IUnknown)) {
        TRACE("(%p)->(IID_IUnknown %p)\n", this, ppv);
        *ppv = static_cast<IMatchCollection2 *>(this);
    } else if (IsEqualGUID(riid, IID_IDispatch)) {
        TRACE("(%p)->(IID_IDispatch %p)\n", this, ppv);
        *ppv = static_cast<IMatchCollection2 *>(this);
    } else if (IsEqualGUID(riid, IID_IMatchCollection2)) {
        TRACE("(%p)->(IID_IMatchCollection2 %p)\n", this, ppv);
        *ppv = static_cast<IMatchCollection2 *>(this);
    } else if (IsEqualGUID(riid, IID_IMatchCollection)) {
        TRACE("(%p)->(IID_IMatchCollection %p)\n", this, ppv);
        *ppv = static_cast<IMatchCollection *>(this);
    } else {
        FIXME("(%p)->(%s %p)\n", this, debugstr_guid(riid), ppv);
        *ppv = NULL;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

ULONG STDMETHODCALLTYPE MatchCollection2::AddRef()
{
    LONG const ref = InterlockedIncrement(&this->ref);
    TRACE("(%p) ref=%ld\n", this, ref);
    return ref;
}

ULONG STDMETHODCALLTYPE MatchCollection2::Release()
{
    LONG const ref = InterlockedDecrement(&this->ref);
    TRACE("(%p) ref=%ld\n", this, ref);
    if (!ref) {
        for (DWORD i = 0; i < matches.size(); i++)
            matches[i]->Release();
        delete this;
    }
    return ref;
}

HRESULT STDMETHODCALLTYPE MatchCollection2::GetTypeInfoCount(UINT *pctinfo)
{
    TRACE("(%p)->(%p)\n", this, pctinfo);
    *pctinfo = 1;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE MatchCollection2::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
    TRACE("(%p)->(%u %lu %p)\n", this, iTInfo, lcid, ppTInfo);
    if (iTInfo != 0)
        return DISP_E_BADINDEX;
    (*ppTInfo = typeinfos[MatchCollection2_tid])->AddRef();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE MatchCollection2::GetIDsOfNames(
    REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
    TRACE("(%p)->(%s %p %u %lu %p)\n", this, debugstr_guid(riid),
          rgszNames, cNames, lcid, rgDispId);

    return typeinfos[MatchCollection2_tid]->GetIDsOfNames(rgszNames, cNames, rgDispId);
}

HRESULT STDMETHODCALLTYPE MatchCollection2::Invoke(DISPID dispIdMember,
        REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams,
        VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
    TRACE("(%p)->(%ld %s %ld %d %p %p %p %p)\n", this, dispIdMember, debugstr_guid(riid),
          lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);

    return typeinfos[MatchCollection2_tid]->Invoke(static_cast<IMatchCollection2 *>(this), dispIdMember, wFlags,
            pDispParams, pVarResult, pExcepInfo, puArgErr);
}

HRESULT STDMETHODCALLTYPE MatchCollection2::get_Item(LONG index, IDispatch **ppMatch)
{
    TRACE("(%p)->()\n", this);

    if (!ppMatch)
        return E_POINTER;

    if (index < 0 || index >= static_cast<LONG>(matches.size()))
        return E_INVALIDARG;

    *ppMatch = matches[index];
    matches[index]->AddRef();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE MatchCollection2::get_Count(LONG *pCount)
{
    TRACE("(%p)->()\n", this);

    if (!pCount)
        return E_POINTER;

    *pCount = matches.size();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE MatchCollection2::get__NewEnum(IUnknown **ppEnum)
{
    TRACE("(%p)->(%p)\n", this, ppEnum);

    if (!ppEnum)
        return E_POINTER;

    return MatchCollectionEnum::create(this, ppEnum);
}

HRESULT MatchCollection2::create(MatchCollection2 **match_collection)
{
    HRESULT hres = init_regexp_typeinfo(MatchCollection2_tid);
    if (FAILED(hres))
        return hres;

    MatchCollection2 *ret = new(std::nothrow) MatchCollection2;
    if (!ret)
        return E_OUTOFMEMORY;

    ret->ref = 1;
    *match_collection = ret;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE RegExp2::QueryInterface(REFIID riid, void **ppv)
{
    if (IsEqualGUID(riid, IID_IUnknown)) {
        TRACE("(%p)->(IID_IUnknown %p)\n", this, ppv);
        *ppv = static_cast<IRegExp2 *>(this);
    } else if (IsEqualGUID(riid, IID_IDispatch)) {
        TRACE("(%p)->(IID_IDispatch %p)\n", this, ppv);
        *ppv = static_cast<IRegExp2 *>(this);
    } else if (IsEqualGUID(riid, IID_IRegExp2)) {
        TRACE("(%p)->(IID_IRegExp2 %p)\n", this, ppv);
        *ppv = static_cast<IRegExp2 *>(this);
    } else if (IsEqualGUID(riid, IID_IRegExp)) {
        TRACE("(%p)->(IID_IRegExp %p)\n", this, ppv);
        *ppv = static_cast<IRegExp *>(this);
    } else {
        FIXME("(%p)->(%s %p)\n", this, debugstr_guid(riid), ppv);
        *ppv = NULL;
        return E_NOINTERFACE;
    }

    AddRef();
    return S_OK;
}

ULONG STDMETHODCALLTYPE RegExp2::AddRef()
{
    LONG const ref = InterlockedIncrement(&this->ref);
    TRACE("(%p) ref=%ld\n", this, ref);
    return ref;
}

ULONG STDMETHODCALLTYPE RegExp2::Release()
{
    LONG const ref = InterlockedDecrement(&this->ref);
    TRACE("(%p) ref=%ld\n", this, ref);
    if (!ref) {
        delete this;
    }
    return ref;
}

HRESULT STDMETHODCALLTYPE RegExp2::GetTypeInfoCount(UINT *pctinfo)
{
    TRACE("(%p)->(%p)\n", this, pctinfo);
    *pctinfo = 1;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE RegExp2::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
    TRACE("(%p)->(%u %lu %p)\n", this, iTInfo, lcid, ppTInfo);
    if (iTInfo != 0)
        return DISP_E_BADINDEX;
    (*ppTInfo = typeinfos[RegExp2_tid])->AddRef();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE RegExp2::GetIDsOfNames(
    REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
    TRACE("(%p)->(%s %p %u %lu %p)\n", this, debugstr_guid(riid),
          rgszNames, cNames, lcid, rgDispId);

    return typeinfos[RegExp2_tid]->GetIDsOfNames(rgszNames, cNames, rgDispId);
}

HRESULT STDMETHODCALLTYPE RegExp2::Invoke(DISPID dispIdMember,
        REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams,
        VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
    TRACE("(%p)->(%ld %s %ld %d %p %p %p %p)\n", this, dispIdMember, debugstr_guid(riid),
          lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);

    return typeinfos[RegExp2_tid]->Invoke(static_cast<RegExp2 *>(this), dispIdMember, wFlags,
                                          pDispParams, pVarResult, pExcepInfo, puArgErr);
}

HRESULT STDMETHODCALLTYPE RegExp2::get_Pattern(BSTR *pPattern)
{
    TRACE("(%p)->(%p)\n", this, pPattern);

    if (!pPattern)
        return E_POINTER;

    *pPattern = SysAllocString(pattern.c_str());
    return *pPattern ? S_OK : E_OUTOFMEMORY;
}

HRESULT STDMETHODCALLTYPE RegExp2::put_Pattern(BSTR pattern)
{
    TRACE("(%p)->(%s)\n", this, wine_dbgstr_w(pattern));

    this->pattern = pattern;
    flags |= REG_PATTERN;

    return S_OK;
}

HRESULT STDMETHODCALLTYPE RegExp2::get_IgnoreCase(VARIANT_BOOL *pIgnoreCase)
{
    TRACE("(%p)->(%p)\n", this, pIgnoreCase);

    if (!pIgnoreCase)
        return E_POINTER;

    *pIgnoreCase = flags & REG_FOLD ? VARIANT_TRUE : VARIANT_FALSE;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE RegExp2::put_IgnoreCase(VARIANT_BOOL ignoreCase)
{
    TRACE("(%p)->(%s)\n", this, ignoreCase ? "true" : "false");

    if (ignoreCase)
        flags |= REG_FOLD;
    else
        flags &= ~REG_FOLD;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE RegExp2::get_Global(VARIANT_BOOL *pGlobal)
{
    TRACE("(%p)->(%p)\n", this, pGlobal);

    if (!pGlobal)
        return E_POINTER;

    *pGlobal = flags & REG_GLOB ? VARIANT_TRUE : VARIANT_FALSE;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE RegExp2::put_Global(VARIANT_BOOL global)
{
    TRACE("(%p)->(%s)\n", this, global ? "true" : "false");

    if (global)
        flags |= REG_GLOB;
    else
        flags &= ~REG_GLOB;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE RegExp2::get_Multiline(VARIANT_BOOL *pMultiline)
{
    TRACE("(%p)->(%p)\n", this, pMultiline);

    if (!pMultiline)
        return E_POINTER;

    *pMultiline = flags & REG_MULTILINE ? VARIANT_TRUE : VARIANT_FALSE;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE RegExp2::put_Multiline(VARIANT_BOOL multiline)
{
    TRACE("(%p)->(%s)\n", this, multiline ? "true" : "false");

    if (multiline)
        flags |= REG_MULTILINE;
    else
        flags &= ~REG_MULTILINE;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE RegExp2::Execute(BSTR source, IDispatch **ppMatches)
{
    RE_PREFIX::wcmatch result;
    MatchCollection2 *match_collection;
    IMatch2 *add = NULL;
    HRESULT hres;

    TRACE("(%p)->(%s %p)\n", this, debugstr_w(source), ppMatches);

    update();

    hres = MatchCollection2::create(&match_collection);
    if (FAILED(hres))
        return hres;

    const WCHAR *cp = source;
    while (RE_PREFIX::regex_search(cp, result, regexp)) {
        const WCHAR *const cq = cp;
        cp += result.position();
        hres = Match2::create(cp - source, result, &add);
        cp += result.length();
        if (FAILED(hres))
            break;
        match_collection->matches.push_back(add);
        if (!*cp || !(flags & REG_GLOB))
            break;
        if (cp == cq) {
            if (pattern.length())
                break;
            ++cp;
        }
    }

    if (FAILED(hres)) {
        match_collection->Release();
        return hres;
    }

    *ppMatches = static_cast<IMatchCollection2 *>(match_collection);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE RegExp2::Test(BSTR source, VARIANT_BOOL *pMatch)
{
    RE_PREFIX::wcmatch result;
    HRESULT hres;

    TRACE("(%p)->(%s %p)\n", this, debugstr_w(source), pMatch);

    update();

    hres = RE_PREFIX::regex_search(source, regexp) ? S_OK : S_FALSE;

    if (hres == S_OK) {
        *pMatch = VARIANT_TRUE;
    } else if (hres == S_FALSE) {
        *pMatch = VARIANT_FALSE;
        hres = S_OK;
    }
    return hres;
}

// class AutoBSTR is modified from code found at
// https://github.com/microsoft/print-oem-samples
// Copyright (c) Microsoft Corporation
// SPDX-License-Identifier: MIT

class AutoBSTR
{
public:
    AutoBSTR(BSTR string = NULL)
        : string(string)
    { }

    ~AutoBSTR() {
        ::SysFreeString(string);
    }

    BSTR* Get() {
        return &string;
    }

    operator BSTR() const {
        return string;
    }

private:
    BSTR string;
};

class StrBuf
    : public ZeroInit<StrBuf>
{
public:
    ~StrBuf();
    BSTR bstr() const;
    BOOL ensure_size(unsigned len);
    HRESULT append(const WCHAR *str, DWORD len);
private:
    WCHAR *buf;
    DWORD size;
    DWORD len;
};

StrBuf::~StrBuf()
{
    CoTaskMemFree(buf);
}

BSTR StrBuf::bstr() const
{
    return SysAllocStringLen(buf, len);
}

BOOL StrBuf::ensure_size(unsigned len)
{
    if (len <= this->size)
        return TRUE;

    DWORD new_size = this->size ? this->size << 1 : 16;
    if (new_size < len)
        new_size = len;
    WCHAR *new_buf = static_cast<WCHAR *>(CoTaskMemRealloc(this->buf, new_size * sizeof(WCHAR)));
    if (!new_buf)
        return FALSE;

    this->buf = new_buf;
    this->size = new_size;
    return TRUE;
}

HRESULT StrBuf::append(const WCHAR *str, DWORD len)
{
    if (!len)
        return S_OK;

    if (!ensure_size(this->len + len))
        return E_OUTOFMEMORY;

    memcpy(this->buf + this->len, str, len * sizeof(WCHAR));
    this->len += len;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE RegExp2::Replace(BSTR source, VARIANT replaceVar, BSTR *pDestString)
{
    StrBuf buf;
    RE_PREFIX::wcmatch result;
    VARIANT strv;
    HRESULT hres;

    TRACE("(%p)->(%s %s %p)\n", this, debugstr_w(source), debugstr_variant(&replaceVar), pDestString);

    update();

    V_VT(&strv) = VT_EMPTY;
    hres = VariantChangeType(&strv, &replaceVar, 0, VT_BSTR);
    if (FAILED(hres))
        return hres;

    AutoBSTR const replace = V_BSTR(&strv);
    size_t const replace_len = SysStringLen(replace);
    size_t const source_len = SysStringLen(source);

    const WCHAR *cp = source;
    while (RE_PREFIX::regex_search(cp, result, regexp)) {
        const WCHAR *const cq = cp;
        hres = buf.append(cp, result.position());
        if (FAILED(hres))
            break;
        cp += result.position() + result.length();
        const WCHAR *prev_ptr = replace;
        while (const WCHAR *ptr = wmemchr(prev_ptr, '$', replace + replace_len - prev_ptr)) {
            hres = buf.append(prev_ptr, ptr - prev_ptr);
            if (FAILED(hres))
                break;
            switch (ptr[1]) {
            case '$':
                hres = buf.append(ptr, 1);
                prev_ptr = ptr + 2;
                break;
            case '&':
                hres = buf.append(cp - result.length(), result.length());
                prev_ptr = ptr + 2;
                break;
            case '`':
                hres = buf.append(source, cp - source - result.length());
                prev_ptr = ptr + 2;
                break;
            case '\'':
                hres = buf.append(cp, source + source_len - cp);
                prev_ptr = ptr + 2;
                break;
            default:
                if (!is_digit(ptr[1])) {
                    hres = buf.append(ptr, 1);
                    prev_ptr = ptr + 1;
                    break;
                }
                size_t paren_count = result.size();
                size_t idx = ptr[1] - '0';
                if (is_digit(ptr[2]) && idx * 10 + (ptr[2] - '0') <= paren_count) {
                    idx = idx * 10 + (ptr[2] - '0');
                    prev_ptr = ptr + 3;
                } else if (idx && idx <= paren_count) {
                    prev_ptr = ptr + 2;
                } else {
                    hres = buf.append(ptr, 1);
                    prev_ptr = ptr + 1;
                    break;
                }

                RE_PREFIX::wcmatch::const_reference sm = result[idx];
                hres = buf.append(sm.first, sm.second - sm.first);
                break;
            }
            if (FAILED(hres))
                break;
        }
        if (FAILED(hres))
            break;
        hres = buf.append(prev_ptr, replace + replace_len - prev_ptr);
        if (FAILED(hres))
            break;
        if (!*cp || !(flags & REG_GLOB))
            break;
        if (cp == cq) {
            if (pattern.length())
                break;
            hres = buf.append(cp++, 1);
            if (FAILED(hres))
                break;
        }
    }

    if (SUCCEEDED(hres)) {
        hres = buf.append(cp, source + source_len - cp);
        if (SUCCEEDED(hres) && (*pDestString = buf.bstr()) == NULL)
            hres = E_OUTOFMEMORY;
    }

    return hres;
}

HRESULT STDMETHODCALLTYPE RegExp2::Replace(BSTR source, BSTR replace, BSTR *pDestString)
{
    VARIANT replaceVar;

    V_VT(&replaceVar) = VT_BSTR;
    V_BSTR(&replaceVar) = replace;
    return Replace(source, replaceVar, pDestString);
}

void RegExp2::update()
{
    if (state != flags) {
        state = (flags &= REG_PATTERN - 1);
        regexp.assign(pattern, static_cast<RE_PREFIX::wregex::flag_type>(flags & (REG_GLOB - 1)));
    }
}

HRESULT RegExp2::create(IDispatch **ret)
{
    HRESULT hres = init_regexp_typeinfo(RegExp2_tid);
    if (FAILED(hres))
        return hres;

    RegExp2 *regexp = new(std::nothrow) RegExp2;
    if (!regexp)
        return E_OUTOFMEMORY;

    regexp->ref = 1;
    regexp->flags = RE_PREFIX::wregex::ECMAScript;

    *ret = static_cast<IRegExp2 *>(regexp);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE RegExp2Factory::QueryInterface(REFIID riid, void **ppv)
{
    if (IsEqualGUID(riid, IID_IUnknown) || IsEqualGUID(riid, IID_IClassFactory)) {
        *ppv = static_cast<IClassFactory *>(this);
    } else {
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    return S_OK;
}

ULONG STDMETHODCALLTYPE RegExp2Factory::AddRef()
{
    return 1;
}

ULONG STDMETHODCALLTYPE RegExp2Factory::Release()
{
    return 1;
}

HRESULT STDMETHODCALLTYPE RegExp2Factory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppv)
{
    TRACE("(%p %s %p)\n", pUnkOuter, debugstr_guid(riid), ppv);

    if (pUnkOuter != NULL)
        return CLASS_E_NOAGGREGATION;

    IDispatch *regexp;
    HRESULT hres = RegExp2::create(&regexp);
    if (FAILED(hres))
        return hres;

    hres = regexp->QueryInterface(riid, ppv);
    regexp->Release();
    return hres;
}

HRESULT STDMETHODCALLTYPE RegExp2Factory::LockServer(BOOL)
{
    return E_NOTIMPL;
}

BOOL APIENTRY DllMain(HANDLE module, DWORD reason, void *)
{
    if (reason == DLL_PROCESS_ATTACH) {
        g_module = reinterpret_cast<HMODULE>(module);
#ifdef _DEBUG
        _strdup("INTENDEDMEMLEAK");
#endif
    }
    return TRUE;
}

STDAPI DllGetClassObject(const CLSID &clsid, const IID &iid, void **ppv)
{
    if (clsid == CLSID_RegExp) {
        return gRegExp2Factory.QueryInterface(iid, ppv);
    }
    return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI DllCanUnloadNow()
{
    return S_FALSE;
}

#define OBJECT_ID L"SRELL.RegExp"
#define OBJECT_DESCRIPTION L"SRELL Regular Expression"
#define OBJECT_VERSION L"1"
#define CLS_ID L"{3f4daca4-81dc-11e1-b0c4-0800200c9a66}"

LPCWSTR const g_RegTable[][3] = {
    { OBJECT_ID, 0, OBJECT_DESCRIPTION },
    { OBJECT_ID L"\\CLSID", 0, CLS_ID },

    { OBJECT_ID L"." OBJECT_VERSION, OBJECT_DESCRIPTION },
    { OBJECT_ID L"." OBJECT_VERSION L"\\CLSID", 0, CLS_ID },

    { L"CLSID\\" CLS_ID, 0, OBJECT_DESCRIPTION },
    { L"CLSID\\" CLS_ID L"\\ProgID", 0, OBJECT_ID L"." OBJECT_VERSION },
    { L"CLSID\\" CLS_ID L"\\VersionIndependentProgID", 0, OBJECT_ID },
    { L"CLSID\\" CLS_ID L"\\InprocServer32", 0, (LPCWSTR)-1 },
};

STDAPI DllUnregisterServer()
{
    HRESULT hr = S_OK;
    for (int i = _countof(g_RegTable) - 1; i >= 0; --i) {
        LPCWSTR pszKeyName = g_RegTable[i][0];

        long err = RegDeleteKeyW(HKEY_CLASSES_ROOT, pszKeyName);
        if (err != ERROR_SUCCESS) {
            hr = S_FALSE;
        }
    }
    return hr;
}

STDAPI DllRegisterServer()
{
    HRESULT hr = S_OK;
    WCHAR szFileName[MAX_PATH];
    GetModuleFileNameW(g_module, szFileName, MAX_PATH);
    for (int i = 0; SUCCEEDED(hr) && i < _countof(g_RegTable); ++i) {
        LPCWSTR pszKeyName = g_RegTable[i][0];
        LPCWSTR pszValueName = g_RegTable[i][1];
        LPCWSTR pszValue = g_RegTable[i][2];

        // -1 is a special marker which says use the DLL file location as the value for the key
        if (pszValue == (LPCWSTR)-1) {
            pszValue = szFileName;
        }

        HKEY hkey;
        long err = RegCreateKeyW(HKEY_CLASSES_ROOT, pszKeyName, &hkey);
        if (err == ERROR_SUCCESS) {
            if (pszValue) {
                err = RegSetValueExW(hkey, pszValueName, 0, REG_SZ, (const BYTE *)pszValue, (wcslen(pszValue) + 1) * sizeof(WCHAR));
            }
            RegCloseKey(hkey);
        }
        if (err != ERROR_SUCCESS) {
            DllUnregisterServer();
            hr = E_FAIL;
        }
    }
    return hr;
}
