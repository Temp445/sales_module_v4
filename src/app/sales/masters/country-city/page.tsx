"use client";

import React, { useState, useEffect } from "react";
import SearchableSelect from "@/components/SearchableSelect";
import { Save, Trash2, XCircle } from "lucide-react";

type TypeMode = "Country" | "City";

export default function CountryCityMaster() {
    const [mode, setMode] = useState<TypeMode>("Country");
    const [countryOptions, setCountryOptions] = useState<{ value: number; label: string }[]>([]);
    const [cityOptions, setCityOptions] = useState<{ value: number; label: string }[]>([]);

    const [selectedCountryCode, setSelectedCountryCode] = useState<number | "">("");
    const [countryName, setCountryName] = useState("");
    
    const [selectedCityCode, setSelectedCityCode] = useState<number | "">("");
    const [cityName, setCityName] = useState("");

    const [isSaving, setIsSaving] = useState(false);
    const [message, setMessage] = useState({ text: "", type: "" });

    // Fetch Countries on load
    const fetchCountries = async () => {
        try {
            const res = await fetch("/api/sales/masters/country-city/fetch-countries");
            const data = await res.json();
            setCountryOptions(data);
        } catch (error) {
            console.error("Failed to fetch countries", error);
        }
    };

    useEffect(() => {
        fetchCountries();
    }, []);

    // Fetch Cities when Country is selected
    useEffect(() => {
        if (selectedCountryCode) {
            const fetchCities = async () => {
                try {
                    const res = await fetch(`/api/sales/masters/country-city/fetch-cities?parentCode=${selectedCountryCode}`);
                    const data = await res.json();
                    setCityOptions(data);
                } catch (error) {
                    console.error("Failed to fetch cities", error);
                }
            };
            fetchCities();
        } else {
            setCityOptions([]);
            setSelectedCityCode("");
            setCityName("");
        }
    }, [selectedCountryCode]);

    const handleModeChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
        setMode(e.target.value as TypeMode);
        handleClear();
    };

    const handleClear = () => {
        setSelectedCountryCode("");
        setCountryName("");
        setSelectedCityCode("");
        setCityName("");
        setMessage({ text: "", type: "" });
    };

    const handleSave = async () => {
        if (mode === "Country" && !countryName.trim()) {
            setMessage({ text: "Enter the Country Name", type: "error" });
            return;
        }
        if (mode === "City") {
            if (!selectedCountryCode) {
                setMessage({ text: "Please Select the Country.", type: "error" });
                return;
            }
            if (!cityName.trim()) {
                setMessage({ text: "Enter the City Name", type: "error" });
                return;
            }
        }

        setIsSaving(true);
        setMessage({ text: "", type: "" });

        const payload = {
            type: mode,
            parentCode: mode === "City" ? selectedCountryCode : 0,
            countryCode: mode === "City" ? selectedCityCode : selectedCountryCode,
            name: mode === "City" ? cityName : countryName
        };

        try {
            const res = await fetch("/api/sales/masters/country-city/save", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload)
            });
            const data = await res.json();

            if (!res.ok || data.error) {
                setMessage({ text: data.error || "Save failed", type: "error" });
            } else {
                setMessage({ text: data.message, type: "success" });
                // Refresh list
                if (mode === "Country") {
                    await fetchCountries();
                    setSelectedCountryCode(data.code);
                } else {
                    // refresh cities
                    const cityRes = await fetch(`/api/sales/masters/country-city/fetch-cities?parentCode=${selectedCountryCode}`);
                    setCityOptions(await cityRes.json());
                    setSelectedCityCode(data.code);
                }
            }
        } catch (error) {
            setMessage({ text: "Network error occurred", type: "error" });
        } finally {
            setIsSaving(false);
        }
    };

    // Calculate Save button state
    const isSaveEnabled = mode === "Country" ? !!countryName.trim() : (!!selectedCountryCode && !!cityName.trim());

    return (
        <div className="p-8 max-w-2xl mx-auto bg-gray-50 min-h-screen">
            <h1 className="text-2xl font-bold text-gray-800 mb-6 border-b pb-2">Country / City Master</h1>
            
            <div className="bg-white p-6 rounded-lg shadow-sm border border-gray-200">
                <div className="flex flex-col gap-y-5">
                    
                    {/* Type Toggle */}
                    <div className="flex items-center">
                        <label className="w-32 font-semibold text-gray-700">Type</label>
                        <select 
                            value={mode}
                            onChange={handleModeChange}
                            className="flex-1 border rounded-md px-3 py-2 outline-none focus:ring-2 focus:ring-blue-500 text-gray-900 bg-white"
                        >
                            <option value="Country">Country</option>
                            <option value="City">City</option>
                        </select>
                    </div>

                    {/* Country Input */}
                    <div className="flex items-center">
                        <label className="w-32 font-semibold text-gray-700">
                            * Country
                        </label>
                        <div className="flex-1">
                            <SearchableSelect
                                options={countryOptions}
                                value={selectedCountryCode || countryName}
                                onChange={(val) => {
                                    if (typeof val === 'number') {
                                        setSelectedCountryCode(val);
                                        const opt = countryOptions.find(o => o.value === val);
                                        if (opt) setCountryName(opt.label);
                                    } else {
                                        setSelectedCountryCode("");
                                        setCountryName(val as string);
                                    }
                                }}
                                onTextChange={(text) => setCountryName(text)}
                                creatable={mode === "Country"}
                                disabled={mode === "City" ? false : false} // always enabled to select parent if City
                                placeholder={mode === "Country" ? "Type new or select Country..." : "Select Country..."}
                            />
                        </div>
                    </div>

                    {/* City Input */}
                    <div className="flex items-center">
                        <label className={`w-32 font-semibold ${mode === 'City' ? 'text-blue-600' : 'text-gray-400'}`}>
                            * City
                        </label>
                        <div className="flex-1">
                            <SearchableSelect
                                options={cityOptions}
                                value={selectedCityCode || cityName}
                                onChange={(val) => {
                                    if (typeof val === 'number') {
                                        setSelectedCityCode(val);
                                        const opt = cityOptions.find(o => o.value === val);
                                        if (opt) setCityName(opt.label);
                                    } else {
                                        setSelectedCityCode("");
                                        setCityName(val as string);
                                    }
                                }}
                                onTextChange={(text) => setCityName(text)}
                                creatable={true}
                                disabled={mode === "Country"}
                                placeholder="Type new or select City..."
                            />
                        </div>
                    </div>
                </div>

                {message.text && (
                    <div className={`mt-4 p-3 rounded text-sm ${message.type === 'error' ? 'bg-red-50 text-red-700' : 'bg-green-50 text-green-700'}`}>
                        {message.text}
                    </div>
                )}

                {/* Actions */}
                <div className="mt-8 flex justify-end gap-3 pt-4 border-t border-gray-100">
                    <button 
                        onClick={handleSave} 
                        disabled={!isSaveEnabled || isSaving}
                        className="flex items-center px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700 disabled:opacity-50 transition-colors"
                    >
                        <Save className="w-4 h-4 mr-2" />
                        Save
                    </button>
                    <button 
                        onClick={handleClear} 
                        className="flex items-center px-4 py-2 bg-gray-100 text-gray-700 rounded hover:bg-gray-200 transition-colors"
                    >
                        <XCircle className="w-4 h-4 mr-2" />
                        Clear
                    </button>
                </div>
            </div>
        </div>
    );
}
