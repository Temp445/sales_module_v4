"use client";

import React, { useState } from "react";
import Link from "next/link";
import { usePathname } from "next/navigation";
import { 
    LayoutDashboard, 
    Map, 
    Settings, 
    LogOut,
    ChevronLeft,
    ChevronRight,
    Building2,
    Users
} from "lucide-react";

const navItems = [
    { name: "Dashboard", href: "/", icon: LayoutDashboard },
    { 
        name: "Sales", 
        icon: Building2,
        isSection: true,
    },
    { name: "Country / City Master", href: "/sales/masters/country-city", icon: Map, indent: true },
    { name: "Document Type Master", href: "/sales/masters/document-type", icon: Map, indent: true },
    { name: "Terms Master", href: "/sales/masters/terms", icon: Map, indent: true },
    { name: "Party Master", href: "/sales/masters/party", icon: Map, indent: true },
    { name: "Customer Item Master", href: "/sales/masters/customer-item", icon: Map, indent: true },
    { name: "Customer Jobwork Master", href: "/sales/masters/jobwork-item", icon: Map, indent: true },
    // { name: "Customers", href: "#", icon: Users, indent: true },
    { name: "Sales Enquiry", href: "/sales/sales-enquiry", icon: Map, indent: true },
    { name: "Sales Quotation", href: "/sales/sales-quotation", icon: Map, indent: true },
    { 
        name: "Settings", 
        icon: Settings,
        isSection: true,
    }
];

export default function Sidebar() {
    const pathname = usePathname();
    const [collapsed, setCollapsed] = useState(false);

    return (
        <div 
            className={`
                h-screen bg-slate-900 border-r border-slate-800 text-slate-300
                flex flex-col transition-all duration-300 ease-in-out relative
                ${collapsed ? "w-20" : "w-64"}
            `}
            style={{
                boxShadow: "10px 0 30px rgba(0,0,0,0.15)",
                backdropFilter: "blur(10px)"
            }}
        >
            {/* Header */}
            <div className="h-16 flex items-center justify-between px-4 border-b border-slate-800 bg-slate-900/50">
                {!collapsed && (
                    <span className="text-xl font-bold bg-gradient-to-r from-blue-400 to-indigo-400 bg-clip-text text-transparent truncate">
                        ERP System
                    </span>
                )}
                {collapsed && (
                    <span className="text-xl font-bold text-blue-400 mx-auto">E</span>
                )}
            </div>

            {/* Collapse Button */}
            <button
                onClick={() => setCollapsed(!collapsed)}
                className="absolute -right-3 top-20 bg-slate-800 text-slate-400 rounded-full p-1 border border-slate-700 hover:text-white hover:bg-slate-700 transition-colors z-10"
            >
                {collapsed ? <ChevronRight size={16} /> : <ChevronLeft size={16} />}
            </button>

            {/* Navigation */}
            <nav className="flex-1 overflow-y-auto py-6 px-3 [&::-webkit-scrollbar]:w-1 [&::-webkit-scrollbar-thumb]:bg-slate-700 [&::-webkit-scrollbar-thumb]:rounded-full">
                <ul className="space-y-1">
                    {navItems.map((item, idx) => {
                        if (item.isSection) {
                            return (
                                <li key={idx} className={`pt-4 pb-2 ${collapsed ? "text-center" : "px-3"}`}>
                                    {collapsed ? (
                                        <div className="mx-auto border-t border-slate-700 w-8" />
                                    ) : (
                                        <span className="text-xs font-semibold text-slate-500 uppercase tracking-wider">
                                            {item.name}
                                        </span>
                                    )}
                                </li>
                            );
                        }

                        const isActive = pathname === item.href;

                        return (
                            <li key={idx}>
                                <Link 
                                    href={item.href!}
                                    className={`
                                        flex items-center rounded-lg px-3 py-2.5 transition-all duration-200 group relative
                                        ${isActive 
                                            ? "bg-blue-600/10 text-blue-400 font-medium" 
                                            : "hover:bg-slate-800/50 hover:text-slate-100"
                                        }
                                        ${item.indent && !collapsed ? "ml-4" : ""}
                                    `}
                                >
                                    <item.icon 
                                        size={20} 
                                        className={`
                                            min-w-[20px] 
                                            ${isActive ? "text-blue-400" : "text-slate-400 group-hover:text-blue-300"}
                                            ${collapsed ? "mx-auto" : "mr-3"}
                                        `}
                                    />
                                    
                                    {!collapsed && (
                                        <span className="truncate">{item.name}</span>
                                    )}
                                    
                                    {isActive && !collapsed && (
                                        <div className="absolute left-0 top-1/2 -translate-y-1/2 w-1 h-5 bg-blue-500 rounded-r-md" />
                                    )}

                                    {/* Tooltip for collapsed state */}
                                    {collapsed && (
                                        <div className="absolute left-full ml-4 px-2 py-1 bg-slate-800 text-xs text-white rounded opacity-0 group-hover:opacity-100 pointer-events-none whitespace-nowrap z-20">
                                            {item.name}
                                        </div>
                                    )}
                                </Link>
                            </li>
                        );
                    })}
                </ul>
            </nav>

            {/* Footer */}
            <div className="p-4 border-t border-slate-800 bg-slate-900/50">
                <Link
                    href="/login"
                    className={`
                        flex items-center text-slate-400 hover:text-red-400 transition-colors group
                        ${collapsed ? "justify-center" : "px-3"}
                    `}
                >
                    <LogOut size={20} className={collapsed ? "" : "mr-3"} />
                    {!collapsed && <span>Logout</span>}
                </Link>
            </div>
        </div>
    );
}
