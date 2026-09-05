"use client";

import React, { useState, useEffect } from "react";
import { Loader2, Save, Trash2, Edit } from "lucide-react";

interface DocumentTypeRecord {
    DocumentType: string;
    Description: string;
    StartingNo: number;
    PertainingTo: string;
}

const PERTAINING_OPTIONS = [
    "Sales Enquiry", "Jobwork Enquiry", "Sales Quotation", "Jobwork Quotation", 
    "Customer PO Entry", "Sales Order", "Jobwork Order", "Schedule Entry", 
    "Inventory Schedule Entry", "Weekly Plan", "Sales Order Amendment", 
    "Jobwork Order Amendment", "Delivery Challan", "Sales Invoice", 
    "Sales Plan Yearly", "Sales Plan Monthly", "Sales Plan Weekly", 
    "Sales Plan Daily", "Contract Review", "Downstream Document"
];

// Map DB code back to UI Text
const CodeToNameMapping: Record<string, string> = {
    "E": "Sales Enquiry", "N": "Jobwork Enquiry", "Q": "Sales Quotation",
    "U": "Jobwork Quotation", "C": "Customer PO Entry", "O": "Sales Order",
    "J": "Jobwork Order", "S": "Schedule Entry", "V": "Inventory Schedule Entry",
    "W": "Weekly Plan", "A": "Sales Order Amendment", "M": "Jobwork Order Amendment",
    "D": "Delivery Challan", "I": "Sales Invoice", "YP": "Sales Plan Yearly",
    "MP": "Sales Plan Monthly", "WP": "Sales Plan Weekly", "DP": "Sales Plan Daily",
    "CR": "Contract Review", "DD": "Downstream Document"
};

export default function DocumentTypeMaster() {
    const [type, setType] = useState("");
    const [desc, setDesc] = useState("");
    const [start, setStart] = useState("");
    const [pertainingTo, setPertainingTo] = useState(PERTAINING_OPTIONS[0]);
    
    const [originalPertainingTo, setOriginalPertainingTo] = useState("");
    const [isUpdate, setIsUpdate] = useState(false);
    
    const [loading, setLoading] = useState(false);
    const [gridData, setGridData] = useState<DocumentTypeRecord[]>([]);
    const [message, setMessage] = useState({ text: "", type: "" });

    useEffect(() => {
        fetchGridData();
    }, []);

    const fetchGridData = async () => {
        try {
            const res = await fetch("/api/sales/masters/document-type/fetch");
            if (res.ok) {
                const data = await res.json();
                setGridData(data);
            }
        } catch (e) {
            console.error("Failed to fetch grid", e);
        }
    };

    const handleSave = async () => {
        setMessage({ text: "", type: "" });
        
        if (!type.trim()) {
            setMessage({ text: "Enter the Document Type", type: "error" });
            return;
        }
        if (!desc.trim()) {
            setMessage({ text: "Enter the Description", type: "error" });
            return;
        }
        if (!start || isNaN(parseInt(start))) {
            setMessage({ text: "Enter a valid Starting Number", type: "error" });
            return;
        }

        setLoading(true);
        try {
            const payload = {
                type,
                desc,
                start,
                pertainingTo,
                isUpdate,
                originalPertainingTo // If we need it later for PK updates
            };

            const res = await fetch("/api/sales/masters/document-type/save", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload)
            });

            const data = await res.json();
            if (res.ok) {
                setMessage({ text: data.message, type: "success" });
                handleClear();
                fetchGridData();
            } else {
                setMessage({ text: data.error, type: "error" });
            }
        } catch (e: any) {
            setMessage({ text: "An error occurred during save.", type: "error" });
        } finally {
            setLoading(false);
        }
    };

    const handleClear = () => {
        setType("");
        setDesc("");
        setStart("");
        setPertainingTo(PERTAINING_OPTIONS[0]);
        setIsUpdate(false);
        setMessage({ text: "", type: "" });
    };

    const handleEdit = (record: DocumentTypeRecord) => {
        setType(record.DocumentType.trim());
        setDesc(record.Description.trim());
        setStart(record.StartingNo.toString());
        
        // Map code back to dropdown string
        const mappedName = CodeToNameMapping[record.PertainingTo.trim()] || PERTAINING_OPTIONS[0];
        setPertainingTo(mappedName);
        setOriginalPertainingTo(mappedName);
        
        setIsUpdate(true);
        setMessage({ text: "", type: "" });
    };

    return (
        <div className="min-h-screen bg-slate-900 text-slate-200 p-8">
            <div className="max-w-5xl mx-auto space-y-6">
                
                <div className="bg-slate-800 border border-slate-700 rounded-xl p-6 shadow-xl relative overflow-hidden">
                    <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-blue-500 to-indigo-500"></div>
                    <h2 className="text-2xl font-semibold mb-6 text-slate-100">Document Type Master</h2>

                    {message.text && (
                        <div className={`p-4 mb-6 rounded-lg font-medium ${message.type === 'error' ? 'bg-red-500/10 text-red-400 border border-red-500/20' : 'bg-green-500/10 text-green-400 border border-green-500/20'}`}>
                            {message.text}
                        </div>
                    )}

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        {/* Left Column */}
                        <div className="space-y-4">
                            <div>
                                <label className="block text-sm font-medium text-slate-400 mb-1">
                                    <span className="text-red-400 mr-1">*</span>Pertaining To
                                </label>
                                <select 
                                    className="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2.5 focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition-all"
                                    value={pertainingTo}
                                    onChange={(e) => setPertainingTo(e.target.value)}
                                >
                                    {PERTAINING_OPTIONS.map(opt => (
                                        <option key={opt} value={opt}>{opt}</option>
                                    ))}
                                </select>
                            </div>

                            <div>
                                <label className="block text-sm font-medium text-slate-400 mb-1">
                                    <span className="text-red-400 mr-1">*</span>Document Type
                                </label>
                                <input 
                                    type="text"
                                    maxLength={3}
                                    className="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2.5 focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition-all uppercase"
                                    value={type}
                                    onChange={(e) => setType(e.target.value.toUpperCase())}
                                    disabled={isUpdate}
                                    placeholder="e.g. INV"
                                />
                                {isUpdate && <p className="text-xs text-slate-500 mt-1">Document Type cannot be changed during update.</p>}
                            </div>
                        </div>

                        {/* Right Column */}
                        <div className="space-y-4">
                            <div>
                                <label className="block text-sm font-medium text-slate-400 mb-1">
                                    <span className="text-red-400 mr-1">*</span>Description
                                </label>
                                <input 
                                    type="text"
                                    maxLength={30}
                                    className="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2.5 focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition-all uppercase"
                                    value={desc}
                                    onChange={(e) => setDesc(e.target.value.toUpperCase())}
                                    placeholder="Enter Description"
                                />
                            </div>

                            <div>
                                <label className="block text-sm font-medium text-slate-400 mb-1">
                                    <span className="text-red-400 mr-1">*</span>Starting No
                                </label>
                                <input 
                                    type="text"
                                    maxLength={5}
                                    className="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2.5 focus:ring-2 focus:ring-blue-500 focus:border-transparent outline-none transition-all"
                                    value={start}
                                    onChange={(e) => {
                                        const val = e.target.value;
                                        if (/^\d*$/.test(val)) setStart(val);
                                    }}
                                    placeholder="0"
                                />
                            </div>
                        </div>
                    </div>

                    <div className="flex items-center justify-end gap-4 mt-8 pt-6 border-t border-slate-700/50">
                        <button 
                            onClick={handleClear}
                            className="px-6 py-2.5 rounded-lg border border-slate-600 text-slate-300 hover:bg-slate-700 hover:text-white transition-all font-medium"
                        >
                            Clear
                        </button>
                        <button 
                            onClick={handleSave}
                            disabled={loading}
                            className="px-6 py-2.5 rounded-lg bg-blue-600 hover:bg-blue-500 text-white transition-all font-medium flex items-center gap-2 shadow-lg shadow-blue-500/20 disabled:opacity-50"
                        >
                            {loading ? <Loader2 className="animate-spin" size={18} /> : <Save size={18} />}
                            {isUpdate ? "Update" : "Save"}
                        </button>
                    </div>
                </div>

                {/* Data Grid */}
                <div className="bg-slate-800 border border-slate-700 rounded-xl overflow-hidden shadow-xl">
                    <div className="px-6 py-4 border-b border-slate-700 bg-slate-800/50 flex justify-between items-center">
                        <h3 className="font-semibold text-slate-200">Existing Document Types</h3>
                        <span className="text-xs bg-slate-700 text-slate-300 px-2 py-1 rounded-full">{gridData.length} Records</span>
                    </div>
                    <div className="overflow-x-auto">
                        <table className="w-full text-left border-collapse">
                            <thead>
                                <tr className="bg-slate-900/50 border-b border-slate-700">
                                    <th className="p-4 font-medium text-slate-400 text-sm">Action</th>
                                    <th className="p-4 font-medium text-slate-400 text-sm">Document Type</th>
                                    <th className="p-4 font-medium text-slate-400 text-sm">Description</th>
                                    <th className="p-4 font-medium text-slate-400 text-sm">Starting No</th>
                                    <th className="p-4 font-medium text-slate-400 text-sm">Pertaining To</th>
                                </tr>
                            </thead>
                            <tbody>
                                {gridData.length === 0 ? (
                                    <tr>
                                        <td colSpan={5} className="p-8 text-center text-slate-500">
                                            No records found.
                                        </td>
                                    </tr>
                                ) : (
                                    gridData.map((row, idx) => (
                                        <tr key={idx} className="border-b border-slate-700/50 hover:bg-slate-700/20 transition-colors group">
                                            <td className="p-4">
                                                <button 
                                                    onClick={() => handleEdit(row)}
                                                    className="p-1.5 text-blue-400 hover:text-blue-300 hover:bg-blue-500/10 rounded transition-colors"
                                                    title="Edit Record"
                                                >
                                                    <Edit size={16} />
                                                </button>
                                            </td>
                                            <td className="p-4 font-medium text-slate-300">{row.DocumentType}</td>
                                            <td className="p-4 text-slate-400">{row.Description}</td>
                                            <td className="p-4 text-slate-400">{row.StartingNo}</td>
                                            <td className="p-4 text-slate-400">
                                                <span className="inline-flex items-center px-2 py-1 rounded-md bg-slate-700/50 text-xs font-medium text-slate-300">
                                                    {CodeToNameMapping[row.PertainingTo.trim()] || row.PertainingTo}
                                                </span>
                                            </td>
                                        </tr>
                                    ))
                                )}
                            </tbody>
                        </table>
                    </div>
                </div>

            </div>
        </div>
    );
}
