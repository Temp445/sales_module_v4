import { NextRequest, NextResponse } from "next/server";
import prisma from "@/lib/prisma";
import { getSession } from "@/lib/session";

export async function GET(req: NextRequest) {
    try {
        const session = await getSession();
        if (!session.isLoggedIn) {
            return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
        }

        const { searchParams } = new URL(req.url);
        const partyCode = searchParams.get("PartyCode");

        if (!partyCode) {
            return NextResponse.json({ error: "PartyCode is required" }, { status: 400 });
        }

        const code = parseInt(partyCode, 10);
        const branchId = session.branchId;
        const compId = session.companyId;
        const compCode = compId; // Legacy assumes first 4 chars of CompID

        // 1. Fetch PartyMaster
        const partyMaster = await prisma.fAS_Party_Master.findFirst({
            where: {
                PartyCode: code,
                Branchid: branchId,
                Compid: compId
            }
        });

        if (!partyMaster) {
            return NextResponse.json({ error: "Party not found" }, { status: 404 });
        }

        // 2. Fetch LedgerMaster (for RegisteredName)
        const ledgerMaster = await prisma.fAS_LedgerMaster.findFirst({
            where: {
                LedgerCode: code,
                CompCode: compCode
            },
            select: {
                RegisteredName: true
            }
        });

        // 3. Fetch PartyAddress
        const partyAddress = await prisma.fAS_PARTY_ADDRESS.findFirst({
            where: {
                PartyCode: code,
                Branchid: branchId,
                Compid: compId
            }
        });

        // Combine into a single payload mapping closely to the UI requirements
        const payload = {
            PartyCode: partyMaster.PartyCode,
            VendorCode: partyMaster.VendorCode,
            SuplierCode: partyMaster.SuplierCode,
            PartyName: partyMaster.PartyName?.trim() || '',
            RegisteredName: ledgerMaster?.RegisteredName?.trim() || '',
            AcType: partyMaster.AcType?.trim() || '',
            FormCode: partyMaster.FormCode,
            PackingCode: partyMaster.PackingCode,
            DeliveryCode: partyMaster.DeliveryCode,
            PaymentCode: partyMaster.PaymentCode,
            InsuranceCode: partyMaster.InsuranceCode,
            ModeCode: partyMaster.ModeCode,
            FreightCode: partyMaster.Freightcode,
            IsSubcontractor: partyMaster.IsSubcontractor,
            AliasName: partyMaster.AliasName?.trim() || '',
            Status: partyMaster.EffStatus ? 0 : 1, // 0 = Active, 1 = Inactive (assumed based on VB6 logic)
            EffDate: partyMaster.EffDate,
            CreditDays: partyMaster.CreditDays,
            CreditDaysPeriod: partyMaster.CreditDaysPeriod?.trim() || '',
            CurrencyCode: partyMaster.CurrencyCode?.trim() || '',
            CustomerFormulaCode: partyMaster.CustomerFormulaCode,
            VendorFormulaCode: partyMaster.VendorFormulaCode,
            TinNo: partyMaster.TinNo?.trim() || '',
            PanNo: partyMaster.PANNo?.trim() || '',

            // Address Details
            Add1: partyAddress?.Add1 || '',
            Add2: partyAddress?.Add2 || '',
            Add3: partyAddress?.Add3 || '',
            Country: partyAddress?.Country?.trim() || '',
            State: partyAddress?.State?.trim() || '',
            City: partyAddress?.City?.trim() || '',
            PinCode: partyAddress?.Pincode?.trim() || '',
            Phone: partyAddress?.Phone?.trim() || '',
            Fax: partyAddress?.Fax?.trim() || '',
            Cellular: partyAddress?.Cellular?.trim() || '',
            Email: partyAddress?.EMail?.trim() || '',
            Cst: partyAddress?.CST?.trim() || '',
            CstDate: partyAddress?.CSTDate || null,
            Lst: partyAddress?.LST?.trim() || '',
            LstDate: partyAddress?.LSTDate || null,
            GSTRegNo: partyAddress?.GSTRegNo?.trim() || '',
            GSTRegDate: partyAddress?.GSTRegDate || null,
            ProvisionalID: partyAddress?.ProvisionalID?.trim() || '',
            UdyamNo: partyAddress?.UdyamNo?.trim() || '', // We might need to split this on frontend if UdyamType is needed
            UdyamDate: partyAddress?.UdyamDate || null,
            AreaCode: partyAddress?.AreaCode?.trim() || '',
            EccNo: partyAddress?.ECCNo?.trim() || '',
            Range: partyAddress?.Range?.trim() || '',
            EDRegNo: partyAddress?.EDRegNo?.trim() || '',
            EDRegDate: partyAddress?.EDRegDate || null,
            ContactPerson: partyAddress?.ContactPerson?.trim() || ''
        };

        return NextResponse.json(payload);
    } catch (error: any) {
        console.error("Fetch Party Error:", error);
        return NextResponse.json({ error: error.message }, { status: 500 });
    }
}
