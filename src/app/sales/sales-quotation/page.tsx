"use client";

import React, { useState, useEffect } from "react";
import SearchableSelect from "@/components/SearchableSelect";

export default function SalesQuotation() {
  const [formData, setFormData] = useState({
    quoType: "",
    quoNo: "",
    quoDate: new Date().toISOString().split("T")[0],
    orderSource: "D", // E = Enquiry, D = Direct
    enqType: "",
    enqNo: "",
    enqDate: "",
    customerCode: "",
    customerName: "",
    validFrom: new Date().toISOString().split("T")[0],
    validTo: new Date().toISOString().split("T")[0],
    formulaCode: "",
    remarks: "",
    formCode: "",
    packingCode: "",
    deliveryCode: "",
    paymentCode: "",
    insuranceCode: "",
    modeCode: "",
    freightCode: "",
    grandTotal: 0
  });

  const [items, setItems] = useState<any[]>([
    { id: 1, itemCode: "", itemDescription: "", uom: "", qty: 0, rate: 0, assessableRate: 0, amount: 0, assGoodsValue: 0 }
  ]);

  // Master Data States
  const [masters, setMasters] = useState({
    quoTypes: [],
    customers: [],
    formulas: [],
    forms: [],
    packings: [],
    deliveries: [],
    payments: [],
    insurances: [],
    modes: [],
    freights: []
  });

  // Item lookup state
  const [itemOptions, setItemOptions] = useState<{value: string; label: string; description: string; rate: number}[]>([]);

  const [loading, setLoading] = useState(true);

  // Fetch Master Data
  useEffect(() => {
    async function loadMasters() {
      try {
        const [
          resDocType, resCust, resFormula, resForm, resPack, resDel, resPay, resIns, resMode, resFreight, resItems
        ] = await Promise.all([
          fetch("/api/sales/masters/document-type/fetch?PertainingTo=Q"),
          fetch("/api/sales/masters/party/search?acType=Debtor"),
          fetch("/api/sales/masters/formula-master/fetch"),
          fetch("/api/sales/masters/terms/fetch?TypeFlg=Form"),
          fetch("/api/sales/masters/terms/fetch?TypeFlg=Packing & Forwarding"),
          fetch("/api/sales/masters/terms/fetch?TypeFlg=Delivery Instruction"),
          fetch("/api/sales/masters/terms/fetch?TypeFlg=Payment Terms"),
          fetch("/api/sales/masters/terms/fetch?TypeFlg=Insurance"),
          fetch("/api/sales/masters/terms/fetch?TypeFlg=Mode"),
          fetch("/api/sales/masters/terms/fetch?TypeFlg=Freight"),
          fetch("/api/sales/masters/item/search")
        ]);

        const dataDoc = await resDocType.json().catch(() => []);
        const dataCust = await resCust.json().catch(() => []);
        const dataFormula = await resFormula.json().catch(() => []);
        const dataItems = await resItems.json().catch(() => []);
        
        const mapDoc = Array.isArray(dataDoc) 
          ? dataDoc.filter((d: any) => String(d.PertainingTo).trim() === 'Q').map((d: any) => ({ value: d.DocumentType, label: d.DocumentType })) 
          : [];
        const mapCust = Array.isArray(dataCust) ? dataCust.map((c: any) => ({ value: c.PartyCode, label: c.PartyName })) : [];
        const mapFormula = Array.isArray(dataFormula) ? dataFormula : [];

        const toOptions = (data: any) => Array.isArray(data) ? data.map((d: any) => ({ value: d.Code, label: d.Description })) : [];

        // Set item options for grid lookup
        if (Array.isArray(dataItems)) {
          setItemOptions(dataItems);
        }

        setMasters({
          quoTypes: mapDoc,
          customers: mapCust,
          formulas: mapFormula,
          forms: toOptions(await resForm.json().catch(() => [])),
          packings: toOptions(await resPack.json().catch(() => [])),
          deliveries: toOptions(await resDel.json().catch(() => [])),
          payments: toOptions(await resPay.json().catch(() => [])),
          insurances: toOptions(await resIns.json().catch(() => [])),
          modes: toOptions(await resMode.json().catch(() => [])),
          freights: toOptions(await resFreight.json().catch(() => []))
        });
      } catch (err) {
        console.error("Failed to load master data", err);
      } finally {
        setLoading(false);
      }
    }
    loadMasters();
  }, []);

  const handleChange = (field: string, val: any) => {
    setFormData((prev) => ({ ...prev, [field]: val }));
  };

  // Auto-fill Terms & Conditions when Customer changes
  useEffect(() => {
    async function loadCustomerDefaults() {
      if (!formData.customerCode) {
        // Clear defaults if customer is deselected
        setFormData(prev => ({
          ...prev,
          formulaCode: "",
          formCode: "",
          packingCode: "",
          deliveryCode: "",
          paymentCode: "",
          insuranceCode: "",
          modeCode: "",
          freightCode: ""
        }));
        return;
      }

      try {
        const res = await fetch(`/api/sales/masters/party/defaults?partyCode=${formData.customerCode}`);
        if (res.ok) {
          const defaults = await res.json();
          setFormData(prev => ({
            ...prev,
            formulaCode: defaults.CustomerFormulaCode ? String(defaults.CustomerFormulaCode) : "",
            formCode: defaults.FormCode ? String(defaults.FormCode) : "",
            packingCode: defaults.PackingCode ? String(defaults.PackingCode) : "",
            deliveryCode: defaults.DeliveryCode ? String(defaults.DeliveryCode) : "",
            paymentCode: defaults.PaymentCode ? String(defaults.PaymentCode) : "",
            insuranceCode: defaults.InsuranceCode ? String(defaults.InsuranceCode) : "",
            modeCode: defaults.ModeCode ? String(defaults.ModeCode) : "",
            freightCode: defaults.Freightcode ? String(defaults.Freightcode) : ""
          }));
        }
      } catch (err) {
        console.error("Failed to fetch customer defaults", err);
      }
    }
    
    loadCustomerDefaults();
  }, [formData.customerCode]);

  const handleItemChange = (index: number, field: string, val: any) => {
    const newItems = [...items];
    newItems[index][field] = val;
    if (field === 'qty' || field === 'rate' || field === 'assessableRate') {
        const qty = parseFloat(newItems[index].qty || 0);
        const rate = parseFloat(newItems[index].rate || 0);
        const assRate = parseFloat(newItems[index].assessableRate || 0);
        newItems[index].amount = qty * rate;          // mGoodsValue = Qty × Rate
        newItems[index].assGoodsValue = qty * assRate; // mAssGoodsValue = Qty × AssRate
    }
    setItems(newItems);
    
    // Update grand total
    const total = newItems.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
    handleChange("grandTotal", total);
  };

  // When item is selected from dropdown, auto-fill description & rate
  const handleItemSelect = (index: number, selectedValue: string | number) => {
    const found = itemOptions.find(o => String(o.value) === String(selectedValue));
    const newItems = [...items];
    newItems[index].itemCode = String(selectedValue);
    if (found) {
      newItems[index].itemDescription = found.description;
      newItems[index].rate = found.rate;
      const qty = parseFloat(newItems[index].qty || 0);
      newItems[index].amount = qty * (found.rate || 0);
      newItems[index].assGoodsValue = qty * parseFloat(newItems[index].assessableRate || 0);
    }
    setItems(newItems);
    const total = newItems.reduce((acc, curr) => acc + (parseFloat(curr.amount) || 0), 0);
    handleChange("grandTotal", total);
  };

  const addItemRow = () => {
    setItems([...items, { id: Date.now(), itemCode: "", itemDescription: "", uom: "", qty: 0, rate: 0, assessableRate: 0, amount: 0, assGoodsValue: 0 }]);
  };

  const handleSave = async () => {
    try {
      const payload = {
        ...formData,
        items: items.filter(i => i.itemCode)
      };
      
      const res = await fetch("/api/sales/transactions/sales-quotation/save", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      const result = await res.json();
      if (res.ok) {
        alert("Sales Quotation Saved Successfully!");
        // Optional clear form
      } else {
        alert("Error: " + result.error);
      }
    } catch (err) {
      alert("Network Error");
    }
  };

  if (loading) return <div className="p-8 text-center text-gray-500">Loading Master Data...</div>;

  return (
    <div className="min-h-screen bg-gray-200 p-6 font-sans text-sm">
      <div className="max-w-7xl mx-auto bg-white rounded-xl shadow-lg border border-gray-100 overflow-hidden">
        
        {/* Header */}
        <div className="bg-gradient-to-r from-blue-700 to-indigo-800 p-5 flex justify-between items-center text-white">
          <h1 className="text-xl font-bold tracking-wide">Sales Quotation</h1>
          <div className="flex space-x-3">
            <button className="px-4 py-2 bg-white/20 hover:bg-white/30 rounded text-sm transition-colors" onClick={() => window.location.reload()}>Clear</button>
            <button className="px-5 py-2 bg-blue-500 hover:bg-blue-400 rounded text-sm font-semibold shadow-md transition-colors" onClick={handleSave}>Save</button>
          </div>
        </div>

        <div className="p-6 space-y-8">
          
          {/* Top Section - Source & Document Info */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6 bg-gray-50 p-5 rounded-lg border border-gray-200">
            {/* Source Options */}
            <div className="flex flex-col justify-center space-y-2">
              <label className="font-semibold text-black">Source</label>
              <div className="flex items-center space-x-4 text-black">
                <label className="flex items-center cursor-pointer">
                  <input type="radio" name="source" checked={formData.orderSource === "E"} onChange={() => handleChange("orderSource", "E")} className="mr-2 text-blue-600 focus:ring-blue-500"/> Enquiry
                </label>
                <label className="flex items-center cursor-pointer">
                  <input type="radio" name="source" checked={formData.orderSource === "D"} onChange={() => handleChange("orderSource", "D")} className="mr-2 text-blue-600 focus:ring-blue-500"/> Direct
                </label>
              </div>
            </div>

            {/* Document Info */}
            <div className="space-y-1">
              <label className="text-black font-medium text-xs uppercase tracking-wider">Quo. Type</label>
              <SearchableSelect options={masters.quoTypes} value={formData.quoType} onChange={(v) => handleChange("quoType", v)} />
            </div>
            <div className="space-y-1">
              <label className="text-black font-medium text-xs uppercase tracking-wider">Quo. No</label>
              <input type="number" value={formData.quoNo} onChange={(e) => handleChange("quoNo", e.target.value)} className="w-full border border-gray-500 rounded-md px-3 py-2 bg-white focus:ring-2 focus:ring-blue-500 outline-none transition-colors text-black placeholder-gray-500" placeholder="Auto/Manual"/>
            </div>
            <div className="space-y-1">
              <label className="text-black font-medium text-xs uppercase tracking-wider">Quo. Date</label>
              <input type="date" value={formData.quoDate} onChange={(e) => handleChange("quoDate", e.target.value)} className="w-full border border-gray-500 rounded-md px-3 py-2 text-black placeholder-gray-500 focus:ring-2 focus:ring-blue-500 outline-none" />
            </div>
          </div>

          {/* Enquiry Info - Show only if Enquiry */}
          {formData.orderSource === "E" && (
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6 bg-blue-50 p-5 rounded-lg border border-blue-100">
              <div className="space-y-1">
                <label className="text-black font-medium text-xs uppercase tracking-wider">Enq. Type</label>
                <input type="text" value={formData.enqType} onChange={(e) => handleChange("enqType", e.target.value)} className="w-full border border-gray-500 rounded-md px-3 py-2 text-black placeholder-gray-500 focus:ring-2 focus:ring-blue-500 outline-none" />
              </div>
              <div className="space-y-1">
                <label className="text-black font-medium text-xs uppercase tracking-wider">Enq. No</label>
                <input type="text" value={formData.enqNo} onChange={(e) => handleChange("enqNo", e.target.value)} className="w-full border border-gray-500 rounded-md px-3 py-2 text-black placeholder-gray-500 focus:ring-2 focus:ring-blue-500 outline-none" />
              </div>
              <div className="space-y-1">
                <label className="text-black font-medium text-xs uppercase tracking-wider">Enq. Date</label>
                <input type="date" value={formData.enqDate} onChange={(e) => handleChange("enqDate", e.target.value)} className="w-full border border-gray-500 rounded-md px-3 py-2 text-black placeholder-gray-500 focus:ring-2 focus:ring-blue-500 outline-none" />
              </div>
            </div>
          )}

          {/* Main Details */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
            <div className="lg:col-span-2 space-y-1">
              <label className="text-black font-medium text-xs uppercase tracking-wider">Customer <span className="text-red-500">*</span></label>
              <SearchableSelect 
                options={masters.customers} 
                value={formData.customerCode} 
                onChange={(v) => handleChange("customerCode", v)}
                onTextChange={(t) => handleChange("customerName", t)}
                placeholder="Select Customer..."
              />
            </div>
            <div className="space-y-1">
              <label className="text-black font-medium text-xs uppercase tracking-wider">Valid From</label>
              <input type="date" value={formData.validFrom} onChange={(e) => handleChange("validFrom", e.target.value)} className="w-full border border-gray-500 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 outline-none text-black placeholder-gray-500" />
            </div>
            <div className="space-y-1">
              <label className="text-black font-medium text-xs uppercase tracking-wider">Valid Till</label>
              <input type="date" value={formData.validTo} onChange={(e) => handleChange("validTo", e.target.value)} className="w-full border border-gray-500 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 outline-none text-black placeholder-gray-500" />
            </div>
            
            <div className="space-y-1 lg:col-span-2">
              <label className="text-black font-medium text-xs uppercase tracking-wider">Formula Name</label>
              <SearchableSelect options={masters.formulas} value={formData.formulaCode} onChange={(v) => handleChange("formulaCode", v)} placeholder="Select Formula..." />
            </div>
            <div className="space-y-1 lg:col-span-2">
              <label className="text-black font-medium text-xs uppercase tracking-wider">Remarks</label>
              <input type="text" value={formData.remarks} onChange={(e) => handleChange("remarks", e.target.value)} className="w-full border border-gray-500 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 outline-none text-black placeholder-gray-500" placeholder="Enter remarks..." />
            </div>
          </div>

          <hr className="border-gray-200" />

          {/* Terms Section */}
          <div>
            <h3 className="text-lg font-semibold text-black mb-4 flex items-center">
              <span className="bg-blue-100 text-blue-800 p-1 rounded mr-2">
                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path></svg>
              </span>
              Terms & Conditions
            </h3>
            <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-4 gap-4">
              <div className="space-y-1">
                <label className="text-black text-xs uppercase tracking-wide">Form</label>
                <SearchableSelect options={masters.forms} value={formData.formCode} onChange={(v) => handleChange("formCode", v)} />
              </div>
              <div className="space-y-1">
                <label className="text-black text-xs uppercase tracking-wide">Packing & Forwarding</label>
                <SearchableSelect options={masters.packings} value={formData.packingCode} onChange={(v) => handleChange("packingCode", v)} />
              </div>
              <div className="space-y-1">
                <label className="text-black text-xs uppercase tracking-wide">Delivery Instruction</label>
                <SearchableSelect options={masters.deliveries} value={formData.deliveryCode} onChange={(v) => handleChange("deliveryCode", v)} />
              </div>
              <div className="space-y-1">
                <label className="text-black text-xs uppercase tracking-wide">Payment Terms</label>
                <SearchableSelect options={masters.payments} value={formData.paymentCode} onChange={(v) => handleChange("paymentCode", v)} />
              </div>
              <div className="space-y-1">
                <label className="text-black text-xs uppercase tracking-wide">Insurance</label>
                <SearchableSelect options={masters.insurances} value={formData.insuranceCode} onChange={(v) => handleChange("insuranceCode", v)} />
              </div>
              <div className="space-y-1">
                <label className="text-black text-xs uppercase tracking-wide">Mode</label>
                <SearchableSelect options={masters.modes} value={formData.modeCode} onChange={(v) => handleChange("modeCode", v)} />
              </div>
              <div className="space-y-1">
                <label className="text-black text-xs uppercase tracking-wide">Freight</label>
                <SearchableSelect options={masters.freights} value={formData.freightCode} onChange={(v) => handleChange("freightCode", v)} />
              </div>
            </div>
          </div>

          <hr className="border-gray-200" />

          {/* Details Grid (spItem) */}
          <div>
             <div className="flex justify-between items-end mb-4">
                <h3 className="text-lg font-semibold text-black">Quotation Items</h3>
                <button onClick={addItemRow} className="px-3 py-1.5 bg-green-50 text-green-700 border border-green-200 hover:bg-green-100 rounded text-sm font-medium transition-colors">
                  + Add Item
                </button>
             </div>
             
             <div className="overflow-x-auto border border-gray-200 rounded-lg shadow-sm">
               <table className="min-w-full divide-y divide-gray-200">
                  <thead className="bg-gray-50">
                    <tr>
                      <th className="px-4 py-3 text-left text-xs font-semibold text-black uppercase tracking-wider">Item Code</th>
                      <th className="px-4 py-3 text-left text-xs font-semibold text-black uppercase tracking-wider w-1/4">Description</th>
                      <th className="px-4 py-3 text-left text-xs font-semibold text-black uppercase tracking-wider">UOM</th>
                      <th className="px-4 py-3 text-right text-xs font-semibold text-black uppercase tracking-wider">Qty</th>
                      <th className="px-4 py-3 text-right text-xs font-semibold text-black uppercase tracking-wider">Rate</th>
                      <th className="px-4 py-3 text-right text-xs font-semibold text-black uppercase tracking-wider">Ass. Rate</th>
                      <th className="px-4 py-3 text-right text-xs font-semibold text-black uppercase tracking-wider">Goods Value</th>
                      <th className="px-4 py-3 text-right text-xs font-semibold text-black uppercase tracking-wider">Ass. Goods Value</th>
                    </tr>
                  </thead>
                  <tbody className="bg-white divide-y divide-gray-200">
                    {items.map((item, idx) => (
                      <tr key={item.id} className="hover:bg-blue-50/50 transition-colors">
                        <td className="px-2 py-2 min-w-[140px]">
                           <SearchableSelect
                             options={itemOptions}
                             value={item.itemCode}
                             onChange={(v) => handleItemSelect(idx, v)}
                             placeholder="Search item..."
                           />
                        </td>
                        <td className="px-2 py-2">
                           <input type="text" value={item.itemDescription} onChange={(e) => handleItemChange(idx, "itemDescription", e.target.value)} className="w-full px-2 py-1.5 border border-gray-500 focus:ring-2 focus:ring-blue-500 outline-none rounded bg-white text-black placeholder-gray-500" />
                        </td>
                        <td className="px-2 py-2">
                           <input type="text" value={item.uom} onChange={(e) => handleItemChange(idx, "uom", e.target.value)} className="w-full px-2 py-1.5 border border-gray-500 focus:ring-2 focus:ring-blue-500 outline-none rounded bg-white text-black placeholder-gray-500" />
                        </td>
                        <td className="px-2 py-2">
                           <input type="number" value={item.qty} onChange={(e) => handleItemChange(idx, "qty", e.target.value)} className="w-full px-2 py-1.5 text-right border border-gray-500 focus:ring-2 focus:ring-blue-500 outline-none rounded bg-white text-black placeholder-gray-500" />
                        </td>
                        <td className="px-2 py-2">
                           <input type="number" value={item.rate} onChange={(e) => handleItemChange(idx, "rate", e.target.value)} className="w-full px-2 py-1.5 text-right border border-gray-500 focus:ring-2 focus:ring-blue-500 outline-none rounded bg-white text-black placeholder-gray-500" />
                        </td>
                        <td className="px-2 py-2">
                           <input type="number" value={item.assessableRate} onChange={(e) => handleItemChange(idx, "assessableRate", e.target.value)} className="w-full px-2 py-1.5 text-right border border-gray-500 focus:ring-2 focus:ring-blue-500 outline-none rounded bg-white text-black placeholder-gray-500" />
                        </td>
                        <td className="px-4 py-2 text-right font-medium text-black bg-gray-50/50">
                           {parseFloat(item.amount || 0).toFixed(2)}
                        </td>
                        <td className="px-4 py-2 text-right font-medium text-black bg-gray-50/50">
                           {parseFloat(item.assGoodsValue || 0).toFixed(2)}
                        </td>
                      </tr>
                    ))}
                  </tbody>
               </table>
             </div>
             
             <div className="flex justify-end mt-4">
                <div className="bg-gray-800 text-white px-6 py-3 rounded-lg shadow-md flex items-center space-x-4">
                  <span className="text-gray-300 text-sm uppercase tracking-wider">Grand Total</span>
                  <span className="text-xl font-bold">{formData.grandTotal.toFixed(2)}</span>
                </div>
             </div>
          </div>

        </div>
      </div>
    </div>
  );
}
