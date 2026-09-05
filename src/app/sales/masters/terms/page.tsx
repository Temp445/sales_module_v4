"use client";

import React, { useState, useEffect } from "react";
import { Loader2, Save, Edit } from "lucide-react";

interface TermRecord {
    TypeFlg: string;
    Code: string;
    Description: string;
    StatusStr: string;
}

const TERM_OPTIONS = [
    "Form", "Packing & Forwarding", "Delivery Instruction", "Insurance", 
    "Freight", "Mode", "Payment Terms", "Contract Review"
];

const STATUS_OPTIONS = ["Active", "Inactive"];

export default function TermsMaster() {
    const [typeFlg, setTypeFlg] = useState(TERM_OPTIONS[0]);
    const [desc, setDesc] = useState("");
    const [status, setStatus] = useState(STATUS_OPTIONS[0]);
    const [code, setCode] = useState("");
    
    const [isUpdate, setIsUpdate] = useState(false);
    
    const [loading, setLoading] = useState(false);
    const [gridData, setGridData] = useState<TermRecord[]>([]);
    const [message, setMessage] = useState({ text: "", type: "" });

    useEffect(() => {
        fetchGridData();
    }, []);

    const fetchGridData = async () => {
        try {
            const res = await fetch("/api/sales/masters/terms/fetch");
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
        
        if (!desc.trim()) {
            setMessage({ text: "Enter the Description", type: "error" });
            return;
        }

        setLoading(true);
        try {
            const payload = {
                typeFlg,
                desc,
                status,
                isUpdate,
                code
            };

            const res = await fetch("/api/sales/masters/terms/save", {
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
        setDesc("");
        setStatus(STATUS_OPTIONS[0]);
        setCode("");
        setIsUpdate(false);
        setMessage({ text: "", type: "" });
    };

    const handleEdit = (record: TermRecord) => {
        setTypeFlg(record.TypeFlg);
        setDesc(record.Description.trim());
        setStatus(record.StatusStr);
        setCode(record.Code.trim());
        
        setIsUpdate(true);
        setMessage({ text: "", type: "" });
    };

    // Filter grid data based on currently selected term type
    const filteredGridData = gridData.filter(r => r.TypeFlg === typeFlg);

    return (
        <div className="min-h-screen bg-slate-900 text-slate-200 p-8">
            <div className="max-w-5xl mx-auto space-y-6">
                
                <div className="bg-slate-800 border border-slate-700 rounded-xl p-6 shadow-xl relative overflow-hidden">
                    <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-emerald-500 to-teal-500"></div>
                    <h2 className="text-2xl font-semibold mb-6 text-slate-100">Terms Master</h2>

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
                                    <span className="text-red-400 mr-1">*</span>Terms
                                </label>
                                <select 
                                    className="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2.5 focus:ring-2 focus:ring-emerald-500 focus:border-transparent outline-none transition-all disabled:opacity-60"
                                    value={typeFlg}
                                    onChange={(e) => {
                                        setTypeFlg(e.target.value);
                                        if (isUpdate) handleClear(); // clearing edit state if they switch terms
                                    }}
                                    disabled={isUpdate}
                                >
                                    {TERM_OPTIONS.map(opt => (
                                        <option key={opt} value={opt}>{opt}</option>
                                    ))}
                                </select>
                            </div>

                            <div>
                                <label className="block text-sm font-medium text-slate-400 mb-1">
                                    Status
                                </label>
                                <select 
                                    className="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2.5 focus:ring-2 focus:ring-emerald-500 focus:border-transparent outline-none transition-all disabled:opacity-60"
                                    value={status}
                                    onChange={(e) => setStatus(e.target.value)}
                                    disabled={!isUpdate} // VB6 Form has this disabled until update
                                >
                                    {STATUS_OPTIONS.map(opt => (
                                        <option key={opt} value={opt}>{opt}</option>
                                    ))}
                                </select>
                                {!isUpdate && <p className="text-xs text-slate-500 mt-1">Status can only be modified for existing terms.</p>}
                            </div>
                        </div>

                        {/* Right Column */}
                        <div className="space-y-4">
                            <div>
                                <label className="block text-sm font-medium text-slate-400 mb-1">
                                    <span className="text-red-400 mr-1">*</span>Description
                                </label>
                                <textarea 
                                    className="w-full bg-slate-900 border border-slate-700 rounded-lg px-4 py-2.5 focus:ring-2 focus:ring-emerald-500 focus:border-transparent outline-none transition-all uppercase min-h-[110px]"
                                    maxLength={100}
                                    value={desc}
                                    onChange={(e) => setDesc(e.target.value.toUpperCase())}
                                    placeholder="Enter Description"
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
                            className="px-6 py-2.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white transition-all font-medium flex items-center gap-2 shadow-lg shadow-emerald-500/20 disabled:opacity-50"
                        >
                            {loading ? <Loader2 className="animate-spin" size={18} /> : <Save size={18} />}
                            {isUpdate ? "Update" : "Save"}
                        </button>
                    </div>
                </div>

                {/* Data Grid */}
                <div className="bg-slate-800 border border-slate-700 rounded-xl overflow-hidden shadow-xl">
                    <div className="px-6 py-4 border-b border-slate-700 bg-slate-800/50 flex justify-between items-center">
                        <h3 className="font-semibold text-slate-200">Existing Terms: <span className="text-emerald-400">{typeFlg}</span></h3>
                        <span className="text-xs bg-slate-700 text-slate-300 px-2 py-1 rounded-full">{filteredGridData.length} Records</span>
                    </div>
                    <div className="overflow-x-auto max-h-[400px]">
                        <table className="w-full text-left border-collapse">
                            <thead>
                                <tr className="bg-slate-900/50 border-b border-slate-700 sticky top-0">
                                    <th className="p-4 font-medium text-slate-400 text-sm w-16">Action</th>
                                    <th className="p-4 font-medium text-slate-400 text-sm w-24">Code</th>
                                    <th className="p-4 font-medium text-slate-400 text-sm">Description</th>
                                    <th className="p-4 font-medium text-slate-400 text-sm w-32">Status</th>
                                </tr>
                            </thead>
                            <tbody>
                                {filteredGridData.length === 0 ? (
                                    <tr>
                                        <td colSpan={4} className="p-8 text-center text-slate-500">
                                            No records found for this term.
                                        </td>
                                    </tr>
                                ) : (
                                    filteredGridData.map((row, idx) => (
                                        <tr key={idx} className="border-b border-slate-700/50 hover:bg-slate-700/20 transition-colors group">
                                            <td className="p-4">
                                                <button 
                                                    onClick={() => handleEdit(row)}
                                                    className="p-1.5 text-emerald-400 hover:text-emerald-300 hover:bg-emerald-500/10 rounded transition-colors"
                                                    title="Edit Record"
                                                >
                                                    <Edit size={16} />
                                                </button>
                                            </td>
                                            <td className="p-4 font-medium text-slate-300">{row.Code}</td>
                                            <td className="p-4 text-slate-400">{row.Description}</td>
                                            <td className="p-4 text-slate-400">
                                                <span className={`inline-flex items-center px-2 py-1 rounded-md text-xs font-medium ${row.StatusStr === 'Active' ? 'bg-emerald-500/10 text-emerald-400' : 'bg-red-500/10 text-red-400'}`}>
                                                    {row.StatusStr}
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
