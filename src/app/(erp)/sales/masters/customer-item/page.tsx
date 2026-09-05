"use client";

import React, { useState, useEffect } from "react";
import SearchableSelect from "@/components/SearchableSelect";
import { Plus, Trash2, Save } from "lucide-react";

interface Option {
    value: string | number;
    label: string;
}

interface CustomerItem {
    itemCode: Option | null;
    description: string;
    model: string;
    drawingIssuedNo: string;
    customerPartNo: string;
    rate: number;
    assessableRate: number;
    validFrom: string;
    validTo: string;
    formulaCode: Option | null;
    formulaName: string;
    qCCheckRequired: "Yes" | "No";
    status: "Active" | "Inactive";
}

export default function CustomerItemMaster() {
    const [customerCode, setCustomerCode] = useState<string | number>("");
    const [defaultFormula, setDefaultFormula] = useState<string | number>("");
    const [items, setItems] = useState<CustomerItem[]>([]);
    
    // Lookups
    const [customerOptions, setCustomerOptions] = useState<Option[]>([]);
    const [formulaOptions, setFormulaOptions] = useState<Option[]>([]);
    const [itemOptions, setItemOptions] = useState<any[]>([]);

    const compId = "07";
    const branchId = "00";
    const userId = "admin";

    useEffect(() => {
        // Fetch Customers
        fetch(`/api/sales/masters/party/search?q=&type=D&compId=${compId}&branchId=${branchId}`)
            .then(res => res.json())
            .then(data => {
                if (Array.isArray(data)) {
                    setCustomerOptions(data.map(c => ({ 
                        value: c.PartyCode, 
                        label: c.PartyName,
                        formulaCode: c.customerFormulaCode
                    })));
                }
            });
            
        // Fetch Formulas
        fetch(`/api/sales/masters/party/customer-formula?compId=${compId}&branchId=${branchId}`)
            .then(res => res.json())
            .then(data => {
                if (Array.isArray(data)) {
                    setFormulaOptions(data.map(f => ({ value: f.FrmID, label: f.FrmLName })));
                }
            });
            // Fetch Items
        fetch(`/api/sales/masters/item/search`)
            .then(res => res.json())
            .then(data => {
                if (Array.isArray(data)) {
                    setItemOptions(data);
                }
            });
    }, []);

    const handleLoadItems = (custCode: string) => {
        if (!custCode) {
            setItems([]);
            return;
        }
        // Fetch items for customer
        fetch(`/api/sales/masters/customer-item/fetch?customerCode=${custCode}`)
            .then(res => res.json())
            .then(data => {
                if (Array.isArray(data)) {
                    setItems(data);
                }
            })
            .catch(err => {
                console.error("Error fetching items:", err);
                alert("Failed to load items.");
            });
    };

    const handleSave = async () => {
        try {
            const res = await fetch("/api/sales/masters/customer-item/save", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    compId,
                    branchId,
                    userId,
                    customerCode,
                    items
                })
            });
            
            if (res.ok) {
                alert("Saved Successfully!");
            } else {
                alert("Failed to save.");
            }
        } catch (error) {
            console.error("Save error:", error);
            alert("Error saving.");
        }
    };

    const addItemRow = () => {
        setItems([
            ...items, 
            {
                itemCode: null,
                description: "",
                model: "",
                drawingIssuedNo: "",
                customerPartNo: "",
                rate: 0,
                assessableRate: 0,
                validFrom: "",
                validTo: "",
                formulaCode: formulaOptions.find(f => f.value === defaultFormula) || null,
                formulaName: "",
                qCCheckRequired: "Yes",
                status: "Active"
            }
        ]);
    };

    const removeItem = (idx: number) => {
        setItems(items.filter((_, i) => i !== idx));
    };

    const updateItem = (idx: number, field: keyof CustomerItem, val: any) => {
        setItems(prev => {
            const newItems = [...prev];
            newItems[idx] = { ...newItems[idx], [field]: val };
            return newItems;
        });
    };

    return (
        <div className="min-h-screen bg-gray-50 p-6 flex flex-col items-center">
            <div className="w-full max-w-7xl bg-white shadow-xl rounded-2xl overflow-hidden border border-gray-100">
                {/* Header */}
                <div className="bg-gradient-to-r from-blue-700 to-indigo-800 px-6 py-4 flex justify-between items-center">
                    <h1 className="text-2xl font-bold text-white tracking-wide">Customer Item Master</h1>
                    <button 
                        onClick={handleSave}
                        className="flex items-center gap-2 bg-emerald-500 hover:bg-emerald-600 text-white px-5 py-2 rounded-lg font-medium transition-all shadow-md active:scale-95"
                    >
                        <Save size={18} /> Save
                    </button>
                </div>

                {/* Top Controls */}
                <div className="p-6 grid grid-cols-1 md:grid-cols-2 gap-8 border-b border-gray-100 bg-gray-50/50">
                    <div>
                        <label className="block text-sm font-semibold text-gray-700 mb-2">Customer</label>
                        <div className="w-full flex items-center gap-2">
                            <div className="flex-grow">
                                <SearchableSelect 
                                    options={customerOptions}
                                    value={customerCode}
                                    onChange={(val) => {
                                        setCustomerCode(val);
                                        const opt: any = customerOptions.find(c => String(c.value) === String(val));
                                        if (opt && opt.formulaCode) {
                                            setDefaultFormula(opt.formulaCode);
                                        } else {
                                            setDefaultFormula(null);
                                        }
                                        handleLoadItems(val);
                                    }}
                                    placeholder="Select Customer..."
                                />
                            </div>
                        </div>
                    </div>
                    <div>
                        <label className="block text-sm font-semibold text-gray-700 mb-2">Default Formula</label>
                        <div className="w-full">
                            <SearchableSelect 
                                options={formulaOptions}
                                value={defaultFormula}
                                onChange={setDefaultFormula}
                                placeholder="Select Formula..."
                            />
                        </div>
                    </div>
                </div>

                {/* Grid Area */}
                <div className="p-6 overflow-x-auto">
                    <div className="flex justify-between items-center mb-4">
                        <h2 className="text-lg font-semibold text-gray-800">Item Details</h2>
                        <button 
                            onClick={addItemRow}
                            className="flex items-center gap-1 bg-blue-50 text-blue-700 hover:bg-blue-100 px-4 py-2 rounded-lg font-medium transition-colors text-sm text-gray-900 font-medium"
                        >
                            <Plus size={16} /> Add Item
                        </button>
                    </div>

                    <table className="w-full text-left border-collapse whitespace-nowrap min-w-max">
                        <thead>
                            <tr className="bg-gray-100 text-gray-600 text-sm uppercase tracking-wider">
                                <th className="p-3 border-b border-gray-200 rounded-tl-lg font-semibold">Item Code</th>
                                <th className="p-3 border-b border-gray-200 font-semibold">Description</th>
                                <th className="p-3 border-b border-gray-200 font-semibold">Model</th>
                                <th className="p-3 border-b border-gray-200 font-semibold">Drawing Issued No</th>
                                <th className="p-3 border-b border-gray-200 font-semibold">Customer Part No</th>
                                <th className="p-3 border-b border-gray-200 font-semibold">Rate</th>
                                <th className="p-3 border-b border-gray-200 font-semibold">Assessable Rate</th>
                                <th className="p-3 border-b border-gray-200 font-semibold">Valid From</th>
                                <th className="p-3 border-b border-gray-200 font-semibold">Valid To</th>
                                <th className="p-3 border-b border-gray-200 font-semibold">Formula Code</th>
                                <th className="p-3 border-b border-gray-200 font-semibold">Formula Name</th>
                                <th className="p-3 border-b border-gray-200 font-semibold">QC Check Required</th>
                                <th className="p-3 border-b border-gray-200 font-semibold">Status</th>
                                <th className="p-3 border-b border-gray-200 rounded-tr-lg font-semibold text-center">Act</th>
                            </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-100">
                            {items.map((item, idx) => (
                                <tr key={idx} className="hover:bg-gray-50/50 transition-colors group">
                                    <td className="p-2 min-w-[200px]">
                                        <SearchableSelect 
                                            options={itemOptions} 
                                            value={item.itemCode?.value || ""}
                                            onChange={(val) => {
                                                const opt = itemOptions.find(i => String(i.value) === String(val));
                                                updateItem(idx, "itemCode", opt ? { value: opt.value, label: opt.value } : { value: val, label: String(val) });
                                                if (opt) {
                                                    updateItem(idx, "description", opt.description);
                                                    updateItem(idx, "model", opt.model);
                                                    updateItem(idx, "rate", opt.rate);
                                                    updateItem(idx, "assessableRate", opt.rate);
                                                    updateItem(idx, "drawingIssuedNo", opt.drawingNo);
                                                }
                                            }}
                                            creatable={true}
                                        />
                                    </td>
                                    <td className="p-2"><input type="text" value={item.description} onChange={e => updateItem(idx, "description", e.target.value)} className="w-full p-2 border border-gray-200 rounded-md text-sm text-gray-900 font-medium" /></td>
                                    <td className="p-2"><input type="text" value={item.model} onChange={e => updateItem(idx, "model", e.target.value)} className="w-full p-2 border border-gray-200 rounded-md text-sm text-gray-900 font-medium" /></td>
                                    <td className="p-2"><input type="text" value={item.drawingIssuedNo} onChange={e => updateItem(idx, "drawingIssuedNo", e.target.value)} className="w-full p-2 border border-gray-200 rounded-md text-sm text-gray-900 font-medium" /></td>
                                    <td className="p-2"><input type="text" value={item.customerPartNo} onChange={e => updateItem(idx, "customerPartNo", e.target.value)} className="w-full p-2 border border-gray-200 rounded-md text-sm text-gray-900 font-medium" /></td>
                                    <td className="p-2"><input type="number" value={item.rate} onChange={e => updateItem(idx, "rate", parseFloat(e.target.value))} className="w-24 p-2 border border-gray-200 rounded-md text-sm text-gray-900 font-medium" /></td>
                                    <td className="p-2"><input type="number" value={item.assessableRate} onChange={e => updateItem(idx, "assessableRate", parseFloat(e.target.value))} className="w-24 p-2 border border-gray-200 rounded-md text-sm text-gray-900 font-medium" /></td>
                                    <td className="p-2"><input type="date" value={item.validFrom} onChange={e => updateItem(idx, "validFrom", e.target.value)} className="w-full p-2 border border-gray-200 rounded-md text-sm text-gray-900 font-medium" /></td>
                                    <td className="p-2"><input type="date" value={item.validTo} onChange={e => updateItem(idx, "validTo", e.target.value)} className="w-full p-2 border border-gray-200 rounded-md text-sm text-gray-900 font-medium" /></td>
                                    <td className="p-2 min-w-[200px]">
                                        <SearchableSelect 
                                            options={formulaOptions}
                                            value={item.formulaCode?.value || ""}
                                            onChange={(val) => {
                                                const opt = formulaOptions.find(f => f.value === val);
                                                updateItem(idx, "formulaCode", opt || null);
                                                updateItem(idx, "formulaName", opt?.label || "");
                                            }}
                                        />
                                    </td>
                                    <td className="p-2"><input type="text" value={item.formulaName} readOnly className="w-full p-2 border border-gray-200 rounded-md text-sm text-gray-900 font-medium bg-gray-100" /></td>
                                    <td className="p-2">
                                        <select value={item.qCCheckRequired} onChange={e => updateItem(idx, "qCCheckRequired", e.target.value)} className="w-full p-2 border border-gray-200 rounded-md text-sm text-gray-900 font-medium bg-white">
                                            <option value="Yes">Yes</option>
                                            <option value="No">No</option>
                                        </select>
                                    </td>
                                    <td className="p-2">
                                        <select value={item.status} onChange={e => updateItem(idx, "status", e.target.value)} className="w-full p-2 border border-gray-200 rounded-md text-sm text-gray-900 font-medium bg-white">
                                            <option value="Active">Active</option>
                                            <option value="Inactive">Inactive</option>
                                        </select>
                                    </td>
                                    <td className="p-2 text-center">
                                        <button onClick={() => removeItem(idx)} className="text-red-400 hover:text-red-600 transition-colors p-1">
                                            <Trash2 size={18} />
                                        </button>
                                    </td>
                                </tr>
                            ))}
                            {items.length === 0 && (
                                <tr>
                                    <td colSpan={12} className="text-center py-10 text-gray-400 italic">No items added yet.</td>
                                </tr>
                            )}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    );
}
