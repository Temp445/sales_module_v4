"use client";

import React, { useState, useEffect, useRef } from "react";
import { ChevronDown, Search } from "lucide-react";

interface Option {
    value: string | number;
    label: string;
}

interface SearchableSelectProps {
    options: Option[];
    value: string | number | "";
    onChange: (val: string | number) => void;
    onTextChange?: (val: string) => void;
    placeholder?: string;
    disabled?: boolean;
    creatable?: boolean;
}

export default function SearchableSelect({
    options,
    value,
    onChange,
    onTextChange,
    placeholder = "Select...",
    disabled = false,
    creatable = false
}: SearchableSelectProps) {
    const [isOpen, setIsOpen] = useState(false);
    const [search, setSearch] = useState("");
    const wrapperRef = useRef<HTMLDivElement>(null);

    // Sync input display with selected value
    const selectedOption = options.find((o) => String(o.value) === String(value));
    useEffect(() => {
        if (!isOpen) {
            setSearch(selectedOption ? selectedOption.label : (creatable && value ? String(value) : ""));
        }
    }, [value, selectedOption, isOpen, creatable]);

    useEffect(() => {
        function handleClickOutside(event: MouseEvent) {
            if (wrapperRef.current && !wrapperRef.current.contains(event.target as Node)) {
                setIsOpen(false);
            }
        }
        document.addEventListener("mousedown", handleClickOutside);
        return () => document.removeEventListener("mousedown", handleClickOutside);
    }, []);

    const filteredOptions = options.filter((o) =>
        o.label.toLowerCase().includes(search.toLowerCase())
    );

    const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        setSearch(e.target.value);
        if (creatable) {
            if (onTextChange) onTextChange(e.target.value);
            // If creatable, typing updates the value as raw text 
            // We'll let the parent handle it via onChange/onTextChange
            onChange(e.target.value);
        }
        if (!isOpen) setIsOpen(true);
    };

    return (
        <div ref={wrapperRef} className="relative w-full">
            <div
                className={`flex items-center border border-gray-500 rounded-md px-3 py-2 bg-white ${
                    disabled ? "opacity-50 cursor-not-allowed" : "cursor-text"
                }`}
                onClick={() => !disabled && setIsOpen(true)}
            >
                <input
                    type="text"
                    className="w-full outline-none bg-transparent text-black placeholder-gray-500"
                    placeholder={placeholder}
                    value={search}
                    onChange={handleInputChange}
                    disabled={disabled}
                    onFocus={() => !disabled && setIsOpen(true)}
                    onKeyDown={(e) => {
                        if (e.key === "Enter" && isOpen) {
                            e.preventDefault();
                            if (creatable && filteredOptions.length === 0) {
                                onChange(search);
                                if (onTextChange) onTextChange(search);
                            } else if (filteredOptions.length > 0) {
                                onChange(filteredOptions[0].value);
                                if (onTextChange) onTextChange(filteredOptions[0].label);
                                setSearch(filteredOptions[0].label);
                            }
                            setIsOpen(false);
                        }
                    }}
                />
                <ChevronDown className="w-4 h-4 text-gray-400 ml-2 cursor-pointer" onClick={(e) => {
                    e.stopPropagation();
                    if (!disabled) setIsOpen(!isOpen);
                }} />
            </div>

            {isOpen && !disabled && (
                <div className="absolute z-10 w-full mt-1 bg-white border rounded-md shadow-lg max-h-60 overflow-y-auto [&::-webkit-scrollbar]:w-2 [&::-webkit-scrollbar-thumb]:bg-gray-300 [&::-webkit-scrollbar-thumb]:rounded-full">
                    {filteredOptions.length > 0 ? (
                        filteredOptions.map((option, idx) => (
                            <div
                                key={`${option.value}-${idx}`}
                                className="px-3 py-2 cursor-pointer hover:bg-blue-50 text-black"
                                onClick={() => {
                                    onChange(option.value);
                                    if (onTextChange) onTextChange(option.label);
                                    setSearch(option.label);
                                    setIsOpen(false);
                                }}
                            >
                                {option.label}
                            </div>
                        ))
                    ) : (
                        <div className="px-3 py-2 text-gray-500">
                            {creatable ? "Press Save to create new" : "No options found"}
                        </div>
                    )}
                </div>
            )}
        </div>
    );
}
