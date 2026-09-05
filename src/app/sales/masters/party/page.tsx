"use client";

import React, { useState, useEffect } from "react";
import { Loader2, Save, Search, PlusCircle, Trash2 } from "lucide-react";
import SearchableSelect from "@/components/SearchableSelect"; // Assuming it exists and supports basic Select UI

type Tab = "general" | "address" | "details" | "terms";

export default function PartyMaster() {
    const [activeTab, setActiveTab] = useState<Tab>("general");
    const [loading, setLoading] = useState(false);
    const [message, setMessage] = useState({ text: "", type: "" });
    const [isUpdate, setIsUpdate] = useState(false);

    // Dropdown Data State
    const [countries, setCountries] = useState<any[]>([]);
    const [states, setStates] = useState<any[]>([]);
    const [cities, setCities] = useState<any[]>([]);
    const [currencies, setCurrencies] = useState<any[]>([]);
    const [custFormulas, setCustFormulas] = useState<any[]>([]);
    const [vendFormulas, setVendFormulas] = useState<any[]>([]);
    const [terms, setTerms] = useState<any[]>([]);

    // Search lookup items
    const [partyList, setPartyList] = useState<any[]>([]);

    // Form State (General)
    const [PartyCode, setPartyCode] = useState(0); // 0 means new
    const [AcType, setAcType] = useState("P"); // Creditor=C, Debtor=D, Party=P
    const [VendorCode, setVendorCode] = useState("");
    const [SuplierCode, setSuplierCode] = useState("");
    const [PartyName, setPartyName] = useState("");
    const [RegisteredName, setRegisteredName] = useState("");
    const [AliasName, setAliasName] = useState("");
    const [Status, setStatus] = useState("0"); // 0=Active, 1=Inactive
    const [IsSubcontractor, setIsSubcontractor] = useState(false);
    const [TinNo, setTinNo] = useState("");
    const [PanNo, setPanNo] = useState("");
    const [EffDate, setEffDate] = useState("");

    // Form State (Address)
    const [Add1, setAdd1] = useState("");
    const [Add2, setAdd2] = useState("");
    const [Add3, setAdd3] = useState("");
    const [Country, setCountry] = useState("");
    const [State, setState] = useState("");
    const [City, setCity] = useState("");
    const [PinCode, setPinCode] = useState("");
    const [Phone, setPhone] = useState("");
    const [Fax, setFax] = useState("");
    const [Cellular, setCellular] = useState("");
    const [Email, setEmail] = useState("");
    const [Cst, setCst] = useState("");
    const [CstDate, setCstDate] = useState("");
    const [Lst, setLst] = useState("");
    const [LstDate, setLstDate] = useState("");

    // Form State (Details & GST)
    const [GSTRegNo, setGSTRegNo] = useState("");
    const [GSTRegDate, setGSTRegDate] = useState("");
    const [ProvisionalID, setProvisionalID] = useState("");
    const [UdyamType, setUdyamType] = useState("");
    const [UdyamNo, setUdyamNo] = useState("");
    const [UdyamDate, setUdyamDate] = useState("");
    const [AreaCode, setAreaCode] = useState("");
    const [EccNo, setEccNo] = useState("");
    const [EDRegNo, setEDRegNo] = useState("");
    const [EDRegDate, setEDRegDate] = useState("");
    const [ContactPerson, setContactPerson] = useState("");

    // Form State (Terms)
    const [FormCode, setFormCode] = useState("0");
    const [PackingCode, setPackingCode] = useState("0");
    const [DeliveryCode, setDeliveryCode] = useState("0");
    const [PaymentCode, setPaymentCode] = useState("0");
    const [InsuranceCode, setInsuranceCode] = useState("0");
    const [ModeCode, setModeCode] = useState("0");
    const [FreightCode, setFreightCode] = useState("0");
    const [CurrencyCode, setCurrencyCode] = useState("");
    const [CustomerFormulaCode, setCustomerFormulaCode] = useState("0");
    const [VendorFormulaCode, setVendorFormulaCode] = useState("0");
    const [CreditDays, setCreditDays] = useState("0");
    const [CreditDaysPeriod, setCreditDaysPeriod] = useState("");

    useEffect(() => {
        loadDropdowns();
    }, []);

    useEffect(() => {
        if (Country && countries.length > 0) {
            fetchStates(Country);
            fetchCities(Country);
        }
    }, [Country, countries]);

    const loadDropdowns = async () => {
        try {
            // Load base dropdowns
            const [curRes, custFormRes, vendFormRes, termsRes, helpersRes, partiesRes] = await Promise.all([
                fetch("/api/sales/masters/party/currencies"),
                fetch("/api/sales/masters/party/customer-formula"),
                fetch("/api/sales/masters/party/vendor-formula"),
                fetch("/api/sales/masters/terms/fetch"),
                fetch("/api/sales/masters/party/helpers?action=country"),
                fetch("/api/sales/masters/party/search")
            ]);

            if (curRes.ok) setCurrencies(await curRes.json());
            if (custFormRes.ok) setCustFormulas(await custFormRes.json());
            if (vendFormRes.ok) setVendFormulas(await vendFormRes.json());
            if (termsRes.ok) setTerms(await termsRes.json());
            if (helpersRes.ok) setCountries(await helpersRes.json());
            if (partiesRes.ok) setPartyList(await partiesRes.json());

        } catch (e) {
            console.error("Failed to load initial dropdowns", e);
        }
    };

    const fetchStates = async (cCode: string) => {
        const c = countries.find(x => x.CountryName?.trim() === cCode?.trim());
        if (!c) return;
        const res = await fetch(`/api/sales/masters/party/helpers?action=state&countryCode=${c.CountryCode}`);
        if (res.ok) setStates(await res.json());
    };

    const fetchCities = async (cCode: string) => {
        const c = countries.find(x => x.CountryName?.trim() === cCode?.trim());
        if (!c) return;
        const res = await fetch(`/api/sales/masters/country-city/fetch-cities?parentCode=${c.CountryCode}`);
        if (res.ok) {
            const data = await res.json();
            setCities(data.map((item: any) => ({ CityCode: item.value, CityName: item.label })));
        }
    };

    const handleClear = () => {
        setIsUpdate(false);
        setPartyCode(0);
        setAcType("P");
        setVendorCode("");
        setSuplierCode("");
        setPartyName("");
        setRegisteredName("");
        setAliasName("");
        setStatus("0");
        setIsSubcontractor(false);
        setTinNo("");
        setPanNo("");
        setEffDate("");
        
        setAdd1(""); setAdd2(""); setAdd3("");
        setCountry(""); setState(""); setCity("");
        setPinCode(""); setPhone(""); setFax(""); setCellular(""); setEmail("");
        setCst(""); setCstDate(""); setLst(""); setLstDate("");
        
        setGSTRegNo(""); setGSTRegDate(""); setProvisionalID("");
        setUdyamType(""); setUdyamNo(""); setUdyamDate("");
        setAreaCode(""); setEccNo(""); setEDRegNo(""); setEDRegDate(""); setContactPerson("");

        setFormCode("0"); setPackingCode("0"); setDeliveryCode("0"); setPaymentCode("0");
        setInsuranceCode("0"); setModeCode("0"); setFreightCode("0"); setCurrencyCode("");
        setCustomerFormulaCode("0"); setVendorFormulaCode("0"); setCreditDays("0"); setCreditDaysPeriod("");

        setMessage({ text: "", type: "" });
        setActiveTab("general");
    };

    const handleSave = async () => {
        if (!PartyName.trim()) {
            setMessage({ text: "Party Name is required", type: "error" });
            return;
        }

        setLoading(true);
        setMessage({ text: "", type: "" });
        
        try {
            // Generate Code if new
            let codeToSave = PartyCode;
            if (!isUpdate && PartyCode === 0) {
                // simple max code generation from list
                const maxCode = partyList.reduce((max, p) => Math.max(max, p.PartyCode), 0);
                codeToSave = maxCode + 1;
            }

            const payload = {
                isUpdate,
                PartyCode: codeToSave,
                AcType, VendorCode, SuplierCode, PartyName, RegisteredName, AliasName, Status,
                IsSubcontractor, TinNo, PanNo, EffDate,
                
                Add1, Add2, Add3, Country, State, City, PinCode, Phone, Fax, Cellular, Email,
                Cst, CstDate, Lst, LstDate,
                
                GSTRegNo, GSTRegDate, ProvisionalID, UdyamType, UdyamNo, UdyamDate,
                AreaCode, EccNo, Range: EccNo, EDRegNo, EDRegDate, ContactPerson,
                
                FormCode, PackingCode, DeliveryCode, PaymentCode, InsuranceCode, ModeCode, FreightCode,
                CurrencyCode, CustomerFormulaCode, VendorFormulaCode, CreditDays, CreditDaysPeriod
            };

            const res = await fetch("/api/sales/masters/party/save", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload)
            });

            const data = await res.json();
            if (res.ok) {
                setMessage({ text: "Party Master saved successfully!", type: "success" });
                
                // Update dropdown search list
                if (!isUpdate) {
                    setPartyList([...partyList, { PartyCode: codeToSave, PartyName }]);
                }
                
                setIsUpdate(true);
                setPartyCode(codeToSave);
            } else {
                setMessage({ text: data.error, type: "error" });
            }
        } catch (e: any) {
            setMessage({ text: "Error saving record", type: "error" });
        } finally {
            setLoading(false);
        }
    };

    const handleFetch = async (partyCode: number) => {
        setLoading(true);
        handleClear();
        try {
            const res = await fetch(`/api/sales/masters/party/fetch?PartyCode=${partyCode}`);
            if (res.ok) {
                const data = await res.json();
                setIsUpdate(true);
                setPartyCode(data.PartyCode);
                setAcType(data.AcType || "P");
                setVendorCode(data.VendorCode || "");
                setSuplierCode(data.SuplierCode || "");
                setPartyName(data.PartyName || "");
                setRegisteredName(data.RegisteredName || "");
                setAliasName(data.AliasName || "");
                setStatus(data.Status?.toString() || "0");
                setIsSubcontractor(data.IsSubcontractor || false);
                setTinNo(data.TinNo || "");
                setPanNo(data.PanNo || "");
                setEffDate(data.EffDate ? data.EffDate.split('T')[0] : "");

                setAdd1(data.Add1 || ""); setAdd2(data.Add2 || ""); setAdd3(data.Add3 || "");
                setCountry(data.Country || ""); setState(data.State || ""); setCity(data.City || "");
                setPinCode(data.PinCode || ""); setPhone(data.Phone || ""); setFax(data.Fax || ""); 
                setCellular(data.Cellular || ""); setEmail(data.Email || "");
                setCst(data.Cst || ""); setCstDate(data.CstDate ? data.CstDate.split('T')[0] : ""); 
                setLst(data.Lst || ""); setLstDate(data.LstDate ? data.LstDate.split('T')[0] : "");

                setGSTRegNo(data.GSTRegNo || ""); setGSTRegDate(data.GSTRegDate ? data.GSTRegDate.split('T')[0] : "");
                setProvisionalID(data.ProvisionalID || "");
                
                // Udyam is often combined, simple un-combine logic if space exists
                if (data.UdyamNo) {
                    const parts = data.UdyamNo.split(' ');
                    if (parts.length > 1) {
                        setUdyamType(parts.pop());
                        setUdyamNo(parts.join(' '));
                    } else {
                        setUdyamNo(data.UdyamNo);
                    }
                }
                setUdyamDate(data.UdyamDate ? data.UdyamDate.split('T')[0] : "");
                
                setAreaCode(data.AreaCode || ""); setEccNo(data.EccNo || ""); 
                setEDRegNo(data.EDRegNo || ""); setEDRegDate(data.EDRegDate ? data.EDRegDate.split('T')[0] : "");
                setContactPerson(data.ContactPerson || "");

                setFormCode(data.FormCode?.toString() || "0");
                setPackingCode(data.PackingCode?.toString() || "0");
                setDeliveryCode(data.DeliveryCode?.toString() || "0");
                setPaymentCode(data.PaymentCode?.toString() || "0");
                setInsuranceCode(data.InsuranceCode?.toString() || "0");
                setModeCode(data.ModeCode?.toString() || "0");
                setFreightCode(data.FreightCode?.toString() || "0");
                setCurrencyCode(data.CurrencyCode || "");
                setCustomerFormulaCode(data.CustomerFormulaCode?.toString() || "0");
                setVendorFormulaCode(data.VendorFormulaCode?.toString() || "0");
                setCreditDays(data.CreditDays?.toString() || "0");
                setCreditDaysPeriod(data.CreditDaysPeriod || "");

            } else {
                setMessage({ text: "Failed to fetch party", type: "error" });
            }
        } catch (e) {
            setMessage({ text: "Network error", type: "error" });
        } finally {
            setLoading(false);
        }
    };

    const getTermsOptions = (flg: string) => terms.filter(t => t.TypeFlg === flg);

    return (
        <div className="min-h-screen bg-slate-900 text-slate-200 p-8">
            <div className="max-w-6xl mx-auto space-y-6">
                
                {/* Header & Lookup */}
                <div className="bg-slate-800 border border-slate-700 rounded-xl p-6 shadow-xl relative overflow-hidden">
                    <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-emerald-500 to-teal-500"></div>
                    <div className="flex justify-between items-start mb-6">
                        <h2 className="text-2xl font-semibold text-slate-100 flex items-center gap-2">
                            Party Master
                            {isUpdate && <span className="text-xs bg-emerald-500/20 text-emerald-400 px-2 py-1 rounded">Edit Mode - Code: {PartyCode}</span>}
                        </h2>
                        
                        <div className="w-72">
                            <label className="text-xs font-medium text-slate-400 mb-1 block">Lookup Existing Party</label>
                            <select 
                                className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-emerald-500 outline-none"
                                onChange={(e) => {
                                    if(e.target.value) handleFetch(parseInt(e.target.value, 10));
                                }}
                                value={isUpdate ? PartyCode : ""}
                            >
                                <option value="">-- Select to Edit --</option>
                                {partyList.map(p => (
                                    <option key={p.PartyCode} value={p.PartyCode}>{p.PartyName}</option>
                                ))}
                            </select>
                        </div>
                    </div>

                    {message.text && (
                        <div className={`p-4 mb-4 rounded-lg font-medium text-sm ${message.type === 'error' ? 'bg-red-500/10 text-red-400 border border-red-500/20' : 'bg-green-500/10 text-green-400 border border-green-500/20'}`}>
                            {message.text}
                        </div>
                    )}
                </div>

                {/* Main Form Container */}
                <div className="bg-slate-800 border border-slate-700 rounded-xl shadow-xl overflow-hidden flex flex-col md:flex-row min-h-[600px]">
                    
                    {/* Vertical Tabs Sidebar */}
                    <div className="w-full md:w-56 bg-slate-800/80 border-r border-slate-700/50 flex flex-col">
                        <div className="p-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Sections</div>
                        <button 
                            onClick={() => setActiveTab("general")}
                            className={`w-full text-left px-6 py-4 text-sm font-medium transition-all border-l-2 ${activeTab === 'general' ? 'border-emerald-500 bg-emerald-500/10 text-emerald-400' : 'border-transparent text-slate-400 hover:bg-slate-700/50 hover:text-slate-200'}`}
                        >
                            General Info
                        </button>
                        <button 
                            onClick={() => setActiveTab("address")}
                            className={`w-full text-left px-6 py-4 text-sm font-medium transition-all border-l-2 ${activeTab === 'address' ? 'border-emerald-500 bg-emerald-500/10 text-emerald-400' : 'border-transparent text-slate-400 hover:bg-slate-700/50 hover:text-slate-200'}`}
                        >
                            Address Details
                        </button>
                        <button 
                            onClick={() => setActiveTab("details")}
                            className={`w-full text-left px-6 py-4 text-sm font-medium transition-all border-l-2 ${activeTab === 'details' ? 'border-emerald-500 bg-emerald-500/10 text-emerald-400' : 'border-transparent text-slate-400 hover:bg-slate-700/50 hover:text-slate-200'}`}
                        >
                            Reg & GST Details
                        </button>
                        <button 
                            onClick={() => setActiveTab("terms")}
                            className={`w-full text-left px-6 py-4 text-sm font-medium transition-all border-l-2 ${activeTab === 'terms' ? 'border-emerald-500 bg-emerald-500/10 text-emerald-400' : 'border-transparent text-slate-400 hover:bg-slate-700/50 hover:text-slate-200'}`}
                        >
                            Terms & Conditions
                        </button>
                    </div>

                    {/* Form Content Area */}
                    <div className="flex-1 p-8 overflow-y-auto">
                        
                        {/* GENERAL TAB */}
                        {activeTab === "general" && (
                            <div className="space-y-6 animate-in fade-in slide-in-from-right-4 duration-300">
                                <h3 className="text-lg font-semibold text-slate-200 border-b border-slate-700 pb-2">General Information</h3>
                                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                                    <div className="col-span-1 md:col-span-2">
                                        <label className="block text-sm font-medium text-slate-400 mb-1"><span className="text-red-400 mr-1">*</span>Party Name (Ledger Name)</label>
                                        <input 
                                            type="text" maxLength={100} value={PartyName} onChange={e => setPartyName(e.target.value.toUpperCase())}
                                            className="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-200 focus:border-emerald-500 outline-none uppercase font-bold text-lg"
                                        />
                                    </div>
                                    <div className="col-span-1 md:col-span-2">
                                        <label className="block text-sm font-medium text-slate-400 mb-1">Registered Name</label>
                                        <input 
                                            type="text" maxLength={100} value={RegisteredName} onChange={e => setRegisteredName(e.target.value.toUpperCase())}
                                            className="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 text-slate-200 focus:border-emerald-500 outline-none uppercase"
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-sm font-medium text-slate-400 mb-1">Account Type</label>
                                        <select value={AcType} onChange={e => setAcType(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 focus:border-emerald-500 outline-none">
                                            <option value="C">Creditor</option>
                                            <option value="D">Debtor</option>
                                            <option value="P">Party</option>
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-sm font-medium text-slate-400 mb-1">Status</label>
                                        <select value={Status} onChange={e => setStatus(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 focus:border-emerald-500 outline-none">
                                            <option value="0">Active</option>
                                            <option value="1">Inactive</option>
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-sm font-medium text-slate-400 mb-1">Vendor Code</label>
                                        <input type="text" value={VendorCode} onChange={e => setVendorCode(e.target.value.toUpperCase())} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 uppercase outline-none focus:border-emerald-500" />
                                    </div>
                                    <div>
                                        <label className="block text-sm font-medium text-slate-400 mb-1">Supplier Code</label>
                                        <input type="text" value={SuplierCode} onChange={e => setSuplierCode(e.target.value.toUpperCase())} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 uppercase outline-none focus:border-emerald-500" />
                                    </div>
                                    <div>
                                        <label className="block text-sm font-medium text-slate-400 mb-1">Alias Name</label>
                                        <input type="text" value={AliasName} onChange={e => setAliasName(e.target.value.toUpperCase())} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2 uppercase outline-none focus:border-emerald-500" />
                                    </div>
                                    <div className="flex items-center gap-2 mt-6">
                                        <input type="checkbox" id="chkSub" checked={IsSubcontractor} onChange={e => setIsSubcontractor(e.target.checked)} className="w-4 h-4 rounded border-slate-700 bg-slate-900 accent-emerald-500" />
                                        <label htmlFor="chkSub" className="text-sm font-medium text-slate-300">Is Subcontractor</label>
                                    </div>
                                </div>
                            </div>
                        )}

                        {/* ADDRESS TAB */}
                        {activeTab === "address" && (
                            <div className="space-y-6 animate-in fade-in slide-in-from-right-4 duration-300">
                                <h3 className="text-lg font-semibold text-slate-200 border-b border-slate-700 pb-2">Postal Address</h3>
                                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                    <div className="col-span-1 md:col-span-2">
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Address Line 1</label>
                                        <input type="text" value={Add1} onChange={e => setAdd1(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" />
                                    </div>
                                    <div className="col-span-1 md:col-span-2">
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Address Line 2</label>
                                        <input type="text" value={Add2} onChange={e => setAdd2(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" />
                                    </div>
                                    <div className="col-span-1 md:col-span-2">
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Address Line 3</label>
                                        <input type="text" value={Add3} onChange={e => setAdd3(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" />
                                    </div>
                                    
                                    {/* Hierarchy Dropdowns */}
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Country</label>
                                        <select value={Country} onChange={e => setCountry(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500">
                                            <option value="">-- Select --</option>
                                            {countries.map(c => <option key={c.CountryCode} value={c.CountryName}>{c.CountryName}</option>)}
                                            {Country && !countries.some(c => c.CountryName?.trim() === Country.trim()) && <option value={Country}>{Country}</option>}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">State</label>
                                        <select value={State} onChange={e => setState(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" disabled={!Country}>
                                            <option value="">-- Select --</option>
                                            {states.map(s => <option key={s.StateCode} value={s.StateName}>{s.StateName}</option>)}
                                            {State && !states.some(s => s.StateName?.trim() === State.trim()) && <option value={State}>{State}</option>}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">City</label>
                                        <select value={City} onChange={e => setCity(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" disabled={!Country}>
                                            <option value="">-- Select --</option>
                                            {cities.map(c => <option key={c.CityCode} value={c.CityName}>{c.CityName}</option>)}
                                            {City && !cities.some(c => c.CityName?.trim() === City.trim()) && <option value={City}>{City}</option>}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Pin Code</label>
                                        <input type="text" value={PinCode} onChange={e => setPinCode(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" />
                                    </div>
                                    
                                    {/* Contact */}
                                    <div className="col-span-1 md:col-span-2 mt-4"><h3 className="text-sm font-semibold text-slate-300 border-b border-slate-700 pb-1">Contact Details</h3></div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Phone</label>
                                        <input type="text" value={Phone} onChange={e => setPhone(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Cellular</label>
                                        <input type="text" value={Cellular} onChange={e => setCellular(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Fax</label>
                                        <input type="text" value={Fax} onChange={e => setFax(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Email</label>
                                        <input type="email" value={Email} onChange={e => setEmail(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" />
                                    </div>

                                    {/* CST / LST (VB6: fraAddress) */}
                                    <div className="col-span-1 md:col-span-2 mt-4"><h3 className="text-sm font-semibold text-slate-300 border-b border-slate-700 pb-1">CST / LST Details</h3></div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">CST No</label>
                                        <input type="text" value={Cst} onChange={e => setCst(e.target.value.toUpperCase())} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500 uppercase" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">CST Date</label>
                                        <input type="date" value={CstDate} onChange={e => setCstDate(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">LST No</label>
                                        <input type="text" value={Lst} onChange={e => setLst(e.target.value.toUpperCase())} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500 uppercase" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">LST Date</label>
                                        <input type="date" value={LstDate} onChange={e => setLstDate(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" />
                                    </div>
                                </div>
                            </div>
                        )}

                        {/* DETAILS TAB */}
                        {activeTab === "details" && (
                            <div className="space-y-6 animate-in fade-in slide-in-from-right-4 duration-300">
                                <h3 className="text-lg font-semibold text-slate-200 border-b border-slate-700 pb-2">Tax & Registration</h3>
                                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                    {/* GST */}
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">GST Reg No</label>
                                        <input type="text" value={GSTRegNo} onChange={e => setGSTRegNo(e.target.value.toUpperCase())} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500 uppercase" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">GST Date</label>
                                        <input type="date" value={GSTRegDate} onChange={e => setGSTRegDate(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Provisional ID</label>
                                        <input type="text" value={ProvisionalID} onChange={e => setProvisionalID(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">PAN No</label>
                                        <input type="text" value={PanNo} onChange={e => setPanNo(e.target.value.toUpperCase())} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500 uppercase" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">TIN No</label>
                                        <input type="text" value={TinNo} onChange={e => setTinNo(e.target.value.toUpperCase())} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500 uppercase" />
                                    </div>

                                    {/* Excise (VB6: fraDetails - txteccno, txtexregno, dpexregdt, txtareacd) */}
                                    <div className="col-span-1 md:col-span-2 mt-4"><h3 className="text-sm font-semibold text-slate-300 border-b border-slate-700 pb-1">Excise Details</h3></div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">ECC No</label>
                                        <input type="text" value={EccNo} onChange={e => setEccNo(e.target.value.toUpperCase())} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500 uppercase" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Area Code</label>
                                        <input type="text" value={AreaCode} onChange={e => setAreaCode(e.target.value.toUpperCase())} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500 uppercase" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Excise Reg. No</label>
                                        <input type="text" value={EDRegNo} onChange={e => setEDRegNo(e.target.value.toUpperCase())} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500 uppercase" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Excise Reg. Date</label>
                                        <input type="date" value={EDRegDate} onChange={e => setEDRegDate(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" />
                                    </div>

                                    {/* Udyam (VB6: txtUdyamType is TextBox, txtUdyamNo, dtpUdyamDate) */}
                                    <div className="col-span-1 md:col-span-2 mt-4"><h3 className="text-sm font-semibold text-slate-300 border-b border-slate-700 pb-1">Udyam Registration</h3></div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Udyam Type</label>
                                        <input type="text" value={UdyamType} onChange={e => setUdyamType(e.target.value)} placeholder="e.g. Micro, Small, Medium" className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Udyam No</label>
                                        <input type="text" value={UdyamNo} onChange={e => setUdyamNo(e.target.value.toUpperCase())} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" />
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Udyam Reg. Date</label>
                                        <input type="date" value={UdyamDate} onChange={e => setUdyamDate(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" />
                                    </div>

                                    {/* Other (VB6: fraDetails - txtContactPerson) */}
                                    <div className="col-span-1 md:col-span-2 mt-4"><h3 className="text-sm font-semibold text-slate-300 border-b border-slate-700 pb-1">Other</h3></div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Contact Person</label>
                                        <input type="text" value={ContactPerson} onChange={e => setContactPerson(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" />
                                    </div>
                                </div>
                            </div>
                        )}

                        {/* TERMS TAB */}
                        {activeTab === "terms" && (
                            <div className="space-y-6 animate-in fade-in slide-in-from-right-4 duration-300">
                                <h3 className="text-lg font-semibold text-slate-200 border-b border-slate-700 pb-2">Terms, Conditions & Formulas</h3>
                                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Customer Formula</label>
                                        <select value={CustomerFormulaCode} onChange={e => setCustomerFormulaCode(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500">
                                            <option value="0">-- None --</option>
                                            {custFormulas.map(f => <option key={f.FrmID} value={f.FrmID}>{f.FrmLName}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Vendor Formula</label>
                                        <select value={VendorFormulaCode} onChange={e => setVendorFormulaCode(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500">
                                            <option value="0">-- None --</option>
                                            {vendFormulas.map(f => <option key={f.FrmID} value={f.FrmID}>{f.FrmLName}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Currency</label>
                                        <select value={CurrencyCode} onChange={e => setCurrencyCode(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500">
                                            <option value="">-- None --</option>
                                            {currencies.map(c => <option key={c.CurrencyCode} value={c.CurrencyCode}>{c.CurrencyName}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Form Type</label>
                                        <select value={FormCode} onChange={e => setFormCode(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500">
                                            <option value="0">-- None --</option>
                                            {getTermsOptions('Form').map(t => <option key={t.Code} value={t.Code}>{t.Description}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Payment Terms</label>
                                        <select value={PaymentCode} onChange={e => setPaymentCode(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500">
                                            <option value="0">-- None --</option>
                                            {getTermsOptions('Payment Terms').map(t => <option key={t.Code} value={t.Code}>{t.Description}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Delivery Instruction</label>
                                        <select value={DeliveryCode} onChange={e => setDeliveryCode(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500">
                                            <option value="0">-- None --</option>
                                            {getTermsOptions('Delivery Instruction').map(t => <option key={t.Code} value={t.Code}>{t.Description}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Mode of Dispatch</label>
                                        <select value={ModeCode} onChange={e => setModeCode(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500">
                                            <option value="0">-- None --</option>
                                            {getTermsOptions('Mode').map(t => <option key={t.Code} value={t.Code}>{t.Description}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Freight</label>
                                        <select value={FreightCode} onChange={e => setFreightCode(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500">
                                            <option value="0">-- None --</option>
                                            {getTermsOptions('Freight').map(t => <option key={t.Code} value={t.Code}>{t.Description}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Insurance</label>
                                        <select value={InsuranceCode} onChange={e => setInsuranceCode(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500">
                                            <option value="0">-- None --</option>
                                            {getTermsOptions('Insurance').map(t => <option key={t.Code} value={t.Code}>{t.Description}</option>)}
                                        </select>
                                    </div>
                                    <div>
                                        <label className="block text-xs font-medium text-slate-400 mb-1">Packing & Forwarding</label>
                                        <select value={PackingCode} onChange={e => setPackingCode(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500">
                                            <option value="0">-- None --</option>
                                            {getTermsOptions('Packing & Forwarding').map(t => <option key={t.Code} value={t.Code}>{t.Description}</option>)}
                                        </select>
                                    </div>
                                    <div className="flex gap-4">
                                        <div className="flex-1">
                                            <label className="block text-xs font-medium text-slate-400 mb-1">Credit Days</label>
                                            <input type="number" value={CreditDays} onChange={e => setCreditDays(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500" />
                                        </div>
                                        <div className="flex-1">
                                            <label className="block text-xs font-medium text-slate-400 mb-1">Period</label>
                                            <select value={CreditDaysPeriod} onChange={e => setCreditDaysPeriod(e.target.value)} className="w-full bg-slate-900 border border-slate-700 rounded-lg px-3 py-2 outline-none focus:border-emerald-500">
                                                <option value="">-- None --</option>
                                                <option value="Delivery">Delivery</option>
                                                <option value="Last Receipt">Last Receipt</option>
                                                <option value="Last RR Approval">Last RR Approval</option>
                                                <option value="Invoice">Invoice</option>
                                            </select>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        )}
                        
                    </div>
                </div>

                {/* Footer Controls */}
                <div className="flex items-center justify-end gap-4 mt-6">
                    <button 
                        onClick={handleClear}
                        className="px-6 py-2.5 rounded-lg border border-slate-600 text-slate-300 hover:bg-slate-700 hover:text-white transition-all font-medium flex items-center gap-2"
                    >
                        <PlusCircle size={18} />
                        New Entry
                    </button>
                    <button 
                        onClick={handleSave}
                        disabled={loading}
                        className="px-8 py-2.5 rounded-lg bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white transition-all font-bold flex items-center gap-2 shadow-lg shadow-emerald-500/25 disabled:opacity-50"
                    >
                        {loading ? <Loader2 className="animate-spin" size={18} /> : <Save size={18} />}
                        {isUpdate ? "Update Party" : "Save New Party"}
                    </button>
                </div>
                
            </div>
        </div>
    );
}
