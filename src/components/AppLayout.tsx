"use client";

import React from "react";
import { usePathname } from "next/navigation";
import Sidebar from "./Sidebar";

export default function AppLayout({ children }: { children: React.ReactNode }) {
    const pathname = usePathname();

    // Do not render the layout for login page
    if (pathname.startsWith("/login")) {
        return <>{children}</>;
    }

    return (
        <div className="flex h-screen overflow-hidden bg-slate-50">
            <Sidebar />
            <main className="flex-1 overflow-auto relative">
                {children}
            </main>
        </div>
    );
}
