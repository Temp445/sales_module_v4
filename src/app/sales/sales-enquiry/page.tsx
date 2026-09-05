"use client";

import React, { useState, useEffect, useCallback } from "react";
import SearchableSelect from "@/components/SearchableSelect";

// --- Types ---
interface EnquiryItem {
    id: number;
    itemType: string;   // "E" = Existing, "N" = New
    itemCode: string;
    itemDesc: string;
    uom: string;
    qty: number | string;
    tarVolme: number | string;   // Target Volume
    volmeper: string;
    rate: number | string;
    assRate: number | string;    // Assessable Rate
    targtDate: string;           // Target Date
    speci: string;               // Specification
    decimals?: number;           // Derived from UOM NoofDecimals
}

interface FormData {
    enquiryType: string;   // cmbEnqType
    enquiryNo: string;     // txtEnqNo (readonly, auto)
    enquiryDate: string;   // dtpEnqDate (readonly, auto)
    referenceNo: string;   // txtRefNo
    referenceDate: string; // dtpRefDate
    customerType: string;  // cmbCustomerType ("E" or "N")
    customerCode: string;  // txtCustomer (SearchableSelect)
    customerName: string;  // derived
    remarks: string;       // txtRemarks
}

const ENQUIRY_TYPES = ["SE", "PE", "CE"]; // Representative VB6 static options from FRX
const CUSTOMER_TYPES = [
    { value: "E", label: "Existing Customer" },
    { value: "N", label: "New Customer" },
];
const ITEM_TYPES = [
    { value: "E", label: "Existing Item" },
    { value: "N", label: "New Item" },
];

const emptyItem = (): EnquiryItem => ({
    id: Date.now(),
    itemType: "E",
    itemCode: "",
    itemDesc: "",
    uom: "",
    qty: "",
    tarVolme: "",
    volmeper: "",
    rate: "",
    assRate: "",
    targtDate: new Date().toISOString().split("T")[0],
    speci: "",
});

const today = () => new Date().toISOString().split("T")[0];

export default function SalesEnquiryPage() {
    const [formData, setFormData] = useState<FormData>({
        enquiryType: "",
        enquiryNo: "",
        enquiryDate: today(),
        referenceNo: "",
        referenceDate: today(),
        customerType: "E",
        customerCode: "",
        customerName: "",
        remarks: "",
    });

    const [items, setItems] = useState<EnquiryItem[]>([emptyItem()]);
    const [customers, setCustomers] = useState<{ value: string | number; label: string }[]>([]);
    const [enquiryTypeOptions, setEnquiryTypeOptions] = useState<{ value: string; label: string }[]>([]);
    const [itemOptions, setItemOptions] = useState<{ value: string; label: string; uom: string; rate: number }[]>([]);
    const [uomOptions, setUomOptions] = useState<{ value: string; label: string; decimals: number }[]>([]);
    const [loading, setLoading] = useState(true);
    const [saving, setSaving] = useState(false);
    const [message, setMessage] = useState<{ type: "success" | "error"; text: string } | null>(null);

    // --- Load master data ---
    useEffect(() => {
        async function loadMasters() {
            try {
        const [resCust, resItems, resDocType, resUoms] = await Promise.all([
                    fetch("/api/sales/enquiry/customers"),
                    fetch("/api/sales/enquiry/items"),
                    fetch("/api/sales/masters/document-type/fetch?PertainingTo=E"),
                    fetch("/api/sales/masters/uom/fetch")
                ]);

                const custData = await resCust.json().catch(() => []);
                const itemData = await resItems.json().catch(() => []);
                const docTypeData = await resDocType.json().catch(() => []);
                const uomData = await resUoms.json().catch(() => []);

                if (Array.isArray(custData)) {
                    setCustomers(
                        custData.map((c: any) => ({
                            value: String(c.CustomerCode ?? c.customerCode ?? ""),
                            label: String(c.PartyName ?? c.partyName ?? ""),
                        }))
                    );
                }

                if (Array.isArray(itemData)) {
                    setItemOptions(
                        itemData.map((i: any) => ({
                            value: String(i.ItemCode ?? ""),
                            label: String(i.ItemDescription ?? ""),
                            uom: String(i.UOMDesc ?? ""),
                            rate: parseFloat(i.Rate ?? 0),
                        }))
                    );
                }

                // Build enquiry type options from Document Type master
                if (Array.isArray(docTypeData)) {
                    const validDocs = docTypeData.filter((d: any) => String(d.PertainingTo).trim() === 'E');
                    setEnquiryTypeOptions(validDocs.map((d: any) => ({ value: d.DocumentType, label: d.DocumentType })));
                }

                // Build UOM options
                if (Array.isArray(uomData)) {
                    setUomOptions(
                        uomData.map((u: any) => ({
                            value: String(u.SDesc ?? "").trim(),
                            label: String(u.SDesc ?? "").trim(),
                            decimals: Number(u.NoofDecimals) || 0
                        })).filter(u => u.value !== "")
                    );
                }
            } catch (err) {
                console.error("Failed to load master data", err);
            } finally {
                setLoading(false);
            }
        }
        loadMasters();
    }, []);

    // --- Generate next enquiry number ---
    useEffect(() => {
        if (!formData.enquiryType) return;
        // Placeholder: enquiry number will be assigned by the SP; show "AUTO" until saved
        setFormData((prev) => ({ ...prev, enquiryNo: "AUTO" }));
    }, [formData.enquiryType]);

    // --- Handlers ---
    const handleChange = useCallback(<K extends keyof FormData>(field: K, val: FormData[K]) => {
        setFormData((prev) => ({ ...prev, [field]: val }));
    }, []);

    const handleCustomerChange = useCallback(
        (val: string | number) => {
            const code = String(val);
            const cust = customers.find((c) => String(c.value) === code);
            setFormData((prev) => ({
                ...prev,
                customerCode: code,
                customerName: cust?.label ?? "",
            }));
        },
        [customers]
    );

    const handleItemChange = useCallback(
        (index: number, field: keyof EnquiryItem, val: any) => {
            setItems((prev) => {
                const updated = [...prev];
                const row = { ...updated[index], [field]: val };

                // Auto-fill UOM and rate when itemCode changes
                if (field === "itemCode") {
                    const value = String(val);
                    // Spread_Clear_Duplicate equivalent
                    if (updated.some((r, i) => i !== index && r.itemCode === value && value !== "")) {
                        setMessage({ type: "error", text: "Duplicate item code not allowed." });
                        row.itemCode = "";
                        return updated;
                    }
                    const found = itemOptions.find((o) => o.value === value);
                    if (found) {
                        row.uom = found.uom;
                        row.rate = found.rate;
                        row.assRate = found.rate;
                        row.itemDesc = found.label;
                        const matchedUom = uomOptions.find(u => u.value === found.uom);
                        if (matchedUom) row.decimals = matchedUom.decimals;
                    }
                }

                if (field === "uom") {
                    const matchedUom = uomOptions.find(u => u.value === val);
                    if (matchedUom) row.decimals = matchedUom.decimals;
                }

                updated[index] = row;
                return updated;
            });
        },
        [itemOptions, uomOptions]
    );

    const addRow = () => setItems((prev) => [...prev, emptyItem()]);

    const removeRow = (index: number) => {
        if (items.length === 1) return;
        setItems((prev) => prev.filter((_, i) => i !== index));
    };

    const handleClear = () => {
        setFormData({
            enquiryType: "",
            enquiryNo: "",
            enquiryDate: today(),
            referenceNo: "",
            referenceDate: today(),
            customerType: "E",
            customerCode: "",
            customerName: "",
            remarks: "",
        });
        setItems([emptyItem()]);
        setMessage(null);
    };

    const handleSave = async () => {
        if (!formData.enquiryType) {
            setMessage({ type: "error", text: "Please select an Enquiry Type." });
            return;
        }
        if (!formData.customerCode) {
            setMessage({ type: "error", text: "Please select a Customer." });
            return;
        }
        const validItems = items.filter((i) => i.itemCode);
        if (validItems.length === 0) {
            setMessage({ type: "error", text: "Please enter at least one item." });
            return;
        }

        setSaving(true);
        setMessage(null);

        try {
            const res = await fetch("/api/sales/enquiry/save", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ header: formData, items: validItems }),
            });

            const data = await res.json();
            if (!res.ok) {
                setMessage({ type: "error", text: data.error || "Save failed." });
            } else {
                setMessage({ type: "success", text: `Sales Enquiry saved successfully! Enquiry No: ${data.enquiryNo ?? ""}` });
                handleClear();
            }
        } catch (err: any) {
            setMessage({ type: "error", text: err.message || "An unexpected error occurred." });
        } finally {
            setSaving(false);
        }
    };

    // --- Styles (inline, minimal) ---
    const labelStyle: React.CSSProperties = { fontWeight: 600, marginBottom: 2, display: "block", fontSize: 12, color: "#374151" };
    const inputStyle: React.CSSProperties = {
        width: "100%", padding: "5px 8px", border: "1px solid #d1d5db",
        borderRadius: 4, fontSize: 13, background: "#fff", boxSizing: "border-box",
        color: "#000",
    };
    const disabledInputStyle: React.CSSProperties = { ...inputStyle, background: "#f3f4f6", color: "#6b7280" };

    if (loading) {
        return (
            <div style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "60vh" }}>
                <div style={{ textAlign: "center" }}>
                    <div className="spinner" style={{ width: 40, height: 40, border: "4px solid #e5e7eb", borderTop: "4px solid #2563eb", borderRadius: "50%", animation: "spin 0.8s linear infinite", margin: "0 auto 12px" }} />
                    <p style={{ color: "#6b7280" }}>Loading Sales Enquiry...</p>
                </div>
            </div>
        );
    }

    return (
        <div style={{ padding: 20, fontFamily: "'Inter', sans-serif", maxWidth: 1100, margin: "0 auto" }}>
            <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>

            {/* Page Title */}
            <div style={{ marginBottom: 16, borderBottom: "2px solid #2563eb", paddingBottom: 8 }}>
                <h1 style={{ margin: 0, fontSize: 20, fontWeight: 700, color: "#1e3a5f" }}>Sales Enquiry</h1>
            </div>

            {/* Message Banner */}
            {message && (
                <div style={{
                    padding: "10px 14px", borderRadius: 6, marginBottom: 14, fontSize: 13,
                    background: message.type === "success" ? "#d1fae5" : "#fee2e2",
                    color: message.type === "success" ? "#065f46" : "#991b1b",
                    border: `1px solid ${message.type === "success" ? "#6ee7b7" : "#fca5a5"}`,
                }}>
                    {message.text}
                </div>
            )}

            {/* === HEADER SECTION === */}
            <div style={{ background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 8, padding: "14px 18px", marginBottom: 16 }}>
                {/* Row 1: EnqType | EnqNo | EnqDate | RefNo | RefDate | CustomerType */}
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr 1fr 1fr", gap: 12, marginBottom: 12 }}>
                    {/* cmbEnqType */}
                    <div>
                        <label htmlFor="enq-type" style={labelStyle}>Enq. Type *</label>
                        <select
                            id="enq-type"
                            value={formData.enquiryType}
                            onChange={(e) => handleChange("enquiryType", e.target.value)}
                            style={inputStyle}
                        >
                            <option value="">-- Select --</option>
                            {enquiryTypeOptions.map((o) => (
                                <option key={o.value} value={o.value}>{o.label}</option>
                            ))}
                        </select>
                    </div>

                    {/* txtEnqNo (Locked=True, auto) */}
                    <div>
                        <label htmlFor="enq-no" style={labelStyle}>Enq. No.</label>
                        <input
                            id="enq-no"
                            type="text"
                            value={formData.enquiryNo}
                            readOnly
                            disabled
                            style={disabledInputStyle}
                            title="Sales Enquiry No (auto-generated on save)"
                        />
                    </div>

                    {/* dtpEnqDate (Locked=True, auto) */}
                    <div>
                        <label htmlFor="enq-date" style={labelStyle}>Enq. Date</label>
                        <input
                            id="enq-date"
                            type="date"
                            value={formData.enquiryDate}
                            readOnly
                            disabled
                            style={disabledInputStyle}
                            title="Sales Enquiry Date (auto)"
                        />
                    </div>

                    {/* txtRefNo */}
                    <div>
                        <label htmlFor="ref-no" style={labelStyle}>Ref. No.</label>
                        <input
                            id="ref-no"
                            type="text"
                            value={formData.referenceNo}
                            onChange={(e) => handleChange("referenceNo", e.target.value)}
                            maxLength={15}
                            style={inputStyle}
                            placeholder="Reference number"
                        />
                    </div>

                    {/* dtpRefDate */}
                    <div>
                        <label htmlFor="ref-date" style={labelStyle}>Ref. Date</label>
                        <input
                            id="ref-date"
                            type="date"
                            value={formData.referenceDate}
                            onChange={(e) => handleChange("referenceDate", e.target.value)}
                            style={inputStyle}
                            title="Reference Date"
                        />
                    </div>

                    {/* cmbCustomerType */}
                    <div>
                        <label htmlFor="cust-type" style={labelStyle}>Customer Type</label>
                        <select
                            id="cust-type"
                            value={formData.customerType}
                            onChange={(e) => handleChange("customerType", e.target.value)}
                            style={inputStyle}
                        >
                            {CUSTOMER_TYPES.map((o) => (
                                <option key={o.value} value={o.value}>{o.label}</option>
                            ))}
                        </select>
                    </div>
                </div>

                {/* Row 2: Customer (wide) | Remarks */}
                <div style={{ display: "grid", gridTemplateColumns: "1.5fr 1fr", gap: 12 }}>
                    {/* txtCustomer - DblClick lookup → SearchableSelect */}
                    <div>
                        <label htmlFor="customer" style={{ ...labelStyle, color: "#dc2626" }}>Customer *</label>
                        <SearchableSelect
                            options={customers}
                            value={formData.customerCode}
                            onChange={handleCustomerChange}
                            placeholder="Search customer..."
                        />
                    </div>

                    {/* txtRemarks */}
                    <div>
                        <label htmlFor="remarks" style={labelStyle}>Remarks</label>
                        <input
                            id="remarks"
                            type="text"
                            value={formData.remarks}
                            onChange={(e) => handleChange("remarks", e.target.value)}
                            maxLength={50}
                            style={inputStyle}
                            placeholder="Enter remarks"
                        />
                    </div>
                </div>
            </div>

            {/* === ITEMS GRID (spItem / AceSpread) === */}
            <div style={{ background: "#fff", border: "1px solid #e2e8f0", borderRadius: 8, overflow: "hidden", marginBottom: 16 }}>
                <div style={{ background: "#1e3a5f", color: "#fff", padding: "8px 14px", fontSize: 13, fontWeight: 600 }}>
                    Item Details
                </div>
                <div style={{ overflowX: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                        <thead>
                            <tr style={{ background: "#f0f4f8" }}>
                                <th style={thStyle}>#</th>
                                <th style={thStyle}>Item Type</th>
                                <th style={{ ...thStyle, minWidth: 160 }}>Item Code *</th>
                                <th style={{ ...thStyle, minWidth: 160 }}>Description</th>
                                <th style={{ ...thStyle, minWidth: 80 }}>UOM</th>
                                <th style={{ ...thStyle, minWidth: 70 }}>Qty</th>
                                <th style={{ ...thStyle, minWidth: 80 }}>Tgt. Volume</th>
                                <th style={{ ...thStyle, minWidth: 80 }}>Volume Per</th>
                                <th style={{ ...thStyle, minWidth: 70 }}>Rate</th>
                                <th style={{ ...thStyle, minWidth: 70 }}>Ass. Rate</th>
                                <th style={{ ...thStyle, minWidth: 110 }}>Target Date</th>
                                <th style={{ ...thStyle, minWidth: 100 }}>Specification</th>
                                <th style={thStyle}>Del</th>
                            </tr>
                        </thead>
                        <tbody>
                            {items.map((item, idx) => (
                                <tr key={item.id} style={{ borderBottom: "1px solid #e5e7eb", background: idx % 2 === 0 ? "#fff" : "#f9fafb" }}>
                                    <td style={tdStyle}>{idx + 1}</td>

                                    {/* ItemType */}
                                    <td style={tdStyle}>
                                        <select
                                            id={`item-type-${idx}`}
                                            value={item.itemType}
                                            onChange={(e) => handleItemChange(idx, "itemType", e.target.value)}
                                            style={{ ...inputStyle, minWidth: 90 }}
                                        >
                                            {ITEM_TYPES.map((o) => (
                                                <option key={o.value} value={o.value}>{o.label}</option>
                                            ))}
                                        </select>
                                    </td>

                                    {/* ItemCode */}
                                    <td style={tdStyle}>
                                        {item.itemType === "N" ? (
                                            <input 
                                                value={item.itemCode}
                                                onChange={(e) => handleItemChange(idx, "itemCode", e.target.value)}
                                                style={inputStyle}
                                                placeholder="Enter code" 
                                                maxLength={20}
                                            />
                                        ) : (
                                            <SearchableSelect
                                                options={itemOptions.map((o) => ({ value: o.value, label: `${o.value} - ${o.label}` }))}
                                                value={item.itemCode}
                                                onChange={(val) => handleItemChange(idx, "itemCode", String(val))}
                                                placeholder="Select item..."
                                            />
                                        )}
                                    </td>

                                    {/* ItemDescription */}
                                    <td style={tdStyle}>
                                        <input
                                            id={`item-desc-${idx}`}
                                            type="text"
                                            value={item.itemDesc}
                                            onChange={(e) => handleItemChange(idx, "itemDesc", e.target.value)}
                                            style={inputStyle}
                                            maxLength={50}
                                        />
                                    </td>

                                    {/* UOM */}
                                    <td style={tdStyle}>
                                        {item.itemType === "N" ? (
                                            <select
                                                id={`item-uom-${idx}`}
                                                value={item.uom}
                                                onChange={(e) => handleItemChange(idx, "uom", e.target.value)}
                                                style={{ ...inputStyle, minWidth: 60 }}
                                            >
                                                <option value="">--</option>
                                                {uomOptions.map(o => (
                                                    <option key={o.value} value={o.value}>{o.label}</option>
                                                ))}
                                            </select>
                                        ) : (
                                            <input
                                                id={`item-uom-${idx}`}
                                                type="text"
                                                value={item.uom}
                                                disabled
                                                style={{ ...disabledInputStyle, minWidth: 60 }}
                                                maxLength={3}
                                            />
                                        )}
                                    </td>

                                    {/* Qty */}
                                    <td style={tdStyle}>
                                        <input
                                            id={`item-qty-${idx}`}
                                            type="number"
                                            value={item.qty}
                                            onChange={(e) => handleItemChange(idx, "qty", e.target.value)}
                                            style={{ ...inputStyle, minWidth: 60 }}
                                            min={0}
                                            step={item.decimals ? Math.pow(10, -item.decimals) : "any"}
                                        />
                                    </td>

                                    {/* Target Volume */}
                                    <td style={tdStyle}>
                                        <input
                                            id={`item-vol-${idx}`}
                                            type="number"
                                            value={item.tarVolme}
                                            onChange={(e) => handleItemChange(idx, "tarVolme", e.target.value)}
                                            style={{ ...inputStyle, minWidth: 70 }}
                                            min={0}
                                            step={item.decimals ? Math.pow(10, -item.decimals) : "any"}
                                        />
                                    </td>

                                    {/* Volume Per */}
                                    <td style={tdStyle}>
                                        <input
                                            id={`item-volper-${idx}`}
                                            type="text"
                                            value={item.volmeper}
                                            onChange={(e) => handleItemChange(idx, "volmeper", e.target.value)}
                                            style={{ ...inputStyle, minWidth: 70 }}
                                            maxLength={15}
                                        />
                                    </td>

                                    {/* Rate */}
                                    <td style={tdStyle}>
                                        <input
                                            id={`item-rate-${idx}`}
                                            type="number"
                                            value={item.rate}
                                            onChange={(e) => handleItemChange(idx, "rate", e.target.value)}
                                            style={{ ...inputStyle, minWidth: 60 }}
                                            min={0}
                                            step="0.01"
                                        />
                                    </td>

                                    {/* Assessable Rate */}
                                    <td style={tdStyle}>
                                        <input
                                            id={`item-assrate-${idx}`}
                                            type="number"
                                            value={item.assRate}
                                            onChange={(e) => handleItemChange(idx, "assRate", e.target.value)}
                                            style={{ ...inputStyle, minWidth: 60 }}
                                            min={0}
                                            step="0.01"
                                        />
                                    </td>

                                    {/* Target Date */}
                                    <td style={tdStyle}>
                                        <input
                                            id={`item-tardate-${idx}`}
                                            type="date"
                                            value={item.targtDate}
                                            onChange={(e) => handleItemChange(idx, "targtDate", e.target.value)}
                                            style={inputStyle}
                                        />
                                    </td>

                                    {/* Specification */}
                                    <td style={tdStyle}>
                                        <input
                                            id={`item-spec-${idx}`}
                                            type="text"
                                            value={item.speci}
                                            onChange={(e) => handleItemChange(idx, "speci", e.target.value)}
                                            onKeyDown={(e) => {
                                                if (e.key === 'Enter') {
                                                    e.preventDefault();
                                                    addRow();
                                                }
                                            }}
                                            style={{ ...inputStyle, minWidth: 100 }}
                                            maxLength={50}
                                        />
                                    </td>

                                    {/* Delete Row */}
                                    <td style={{ ...tdStyle, textAlign: "center" }}>
                                        <button
                                            id={`del-row-${idx}`}
                                            onClick={() => removeRow(idx)}
                                            disabled={items.length === 1}
                                            style={{
                                                background: "none", border: "none", cursor: items.length === 1 ? "not-allowed" : "pointer",
                                                color: items.length === 1 ? "#d1d5db" : "#ef4444", fontSize: 16, padding: "2px 6px",
                                            }}
                                            title="Remove row"
                                        >
                                            ✕
                                        </button>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>

                {/* Add Row */}
                <div style={{ padding: "8px 14px" }}>
                    <button
                        id="add-item-row"
                        onClick={addRow}
                        style={{
                            background: "#e0f2fe", border: "1px solid #7dd3fc", color: "#0369a1",
                            borderRadius: 4, padding: "5px 14px", fontSize: 12, cursor: "pointer", fontWeight: 600,
                        }}
                    >
                        + Add Row
                    </button>
                </div>
            </div>

            {/* === ACTION BUTTONS === */}
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
                <button
                    id="btn-save"
                    onClick={handleSave}
                    disabled={saving}
                    style={{
                        background: saving ? "#93c5fd" : "#2563eb",
                        color: "#fff", border: "none", borderRadius: 6,
                        padding: "9px 28px", fontSize: 14, fontWeight: 700,
                        cursor: saving ? "not-allowed" : "pointer",
                        boxShadow: "0 1px 3px rgba(0,0,0,0.15)",
                    }}
                >
                    {saving ? "Saving..." : "💾 Save"}
                </button>

                <button
                    id="btn-clear"
                    onClick={handleClear}
                    disabled={saving}
                    style={{
                        background: "#f3f4f6", color: "#374151", border: "1px solid #d1d5db",
                        borderRadius: 6, padding: "9px 22px", fontSize: 14, fontWeight: 600,
                        cursor: saving ? "not-allowed" : "pointer",
                    }}
                >
                    Clear
                </button>

                <button
                    id="btn-exit"
                    onClick={() => window.history.back()}
                    style={{
                        background: "#fef3c7", color: "#92400e", border: "1px solid #fcd34d",
                        borderRadius: 6, padding: "9px 22px", fontSize: 14, fontWeight: 600,
                        cursor: "pointer",
                    }}
                >
                    Exit
                </button>
            </div>
        </div>
    );
}

// --- Table cell styles ---
const thStyle: React.CSSProperties = {
    padding: "7px 8px", textAlign: "left", fontWeight: 700,
    fontSize: 11, color: "#374151", borderBottom: "2px solid #d1d5db",
    whiteSpace: "nowrap",
};

const tdStyle: React.CSSProperties = {
    padding: "4px 6px", verticalAlign: "middle",
};