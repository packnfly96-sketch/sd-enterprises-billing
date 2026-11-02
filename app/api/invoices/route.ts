import { NextResponse } from "next/server";
import connectToDatabase from "@/lib/mongodb";
import Invoice from "@/models/Invoice";

function isInvoicePayload(value: unknown): value is Record<string, unknown> {
  if (typeof value !== "object" || value === null) return false;
  const invoice = value as Record<string, unknown>;

  return (
    typeof invoice.invoiceNo === "string" &&
    typeof invoice.invoiceDate === "string" &&
    typeof invoice.dueDate === "string" &&
    typeof invoice.customerName === "string" &&
    typeof invoice.customerState === "string" &&
    (invoice.invoiceType === "GST" || invoice.invoiceType === "Non-GST") &&
    (invoice.gstType === "IGST" || invoice.gstType === "CGST+SGST") &&
    typeof invoice.taxableValue === "number" &&
    typeof invoice.cgst === "number" &&
    typeof invoice.sgst === "number" &&
    typeof invoice.igst === "number" &&
    typeof invoice.totalAmount === "number" &&
    typeof invoice.status === "string" &&
    typeof invoice.paymentStatus === "string" &&
    Array.isArray(invoice.consignments) &&
    typeof invoice.customer === "object" &&
    invoice.customer !== null
  );
}

export async function GET() {
  try {
    await connectToDatabase();
    const invoices = await Invoice.find().sort({ createdAt: -1 }).lean();
    return NextResponse.json(invoices);
  } catch (error) {
    console.error("Could not fetch invoices from MongoDB.", error);
    return NextResponse.json({ error: "Could not fetch invoices." }, { status: 500 });
  }
}

export async function POST(request: Request) {
  try {
    const invoiceData: unknown = await request.json();
    if (!isInvoicePayload(invoiceData)) {
      return NextResponse.json({ error: "Invalid invoice data." }, { status: 400 });
    }

    await connectToDatabase();
    const invoice = await Invoice.create(invoiceData);
    return NextResponse.json(invoice.toObject(), { status: 201 });
  } catch (error) {
    console.error("Could not create invoice in MongoDB.", error);
    if (error instanceof SyntaxError) {
      return NextResponse.json({ error: "Invalid JSON request body." }, { status: 400 });
    }
    if (error instanceof Error && error.name === "ValidationError") {
      return NextResponse.json({ error: error.message }, { status: 400 });
    }
    return NextResponse.json({ error: "Could not create invoice." }, { status: 500 });
  }
}
