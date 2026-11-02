import { Types } from "mongoose";
import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import connectToDatabase from "@/lib/mongodb";
import Invoice from "@/models/Invoice";

type RouteContext = { params: { id: string } };

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
    invoice.customer !== null &&
    (invoice.createdBy === undefined || typeof invoice.createdBy === "string")
  );
}

export async function PUT(request: Request, { params }: RouteContext) {
  const sessionUser = getSessionUser(request);
  if (!sessionUser) {
    return NextResponse.json({ error: "Authentication required." }, { status: 401 });
  }
  if (!Types.ObjectId.isValid(params.id)) {
    return NextResponse.json({ error: "Invalid invoice ID." }, { status: 400 });
  }

  try {
    const invoiceData: unknown = await request.json();
    if (!isInvoicePayload(invoiceData)) {
      return NextResponse.json({ error: "Invalid invoice data." }, { status: 400 });
    }

    await connectToDatabase();
    const { createdBy: _untrustedCreatedBy, ...editableInvoice } = invoiceData;
    const filter =
      sessionUser.role === "admin"
        ? { _id: params.id }
        : { _id: params.id, createdBy: sessionUser.email };
    const invoice = await Invoice.findOneAndUpdate(filter, editableInvoice, {
      new: true,
      runValidators: true,
    });
    if (!invoice) {
      return NextResponse.json({ error: "Invoice not found." }, { status: 404 });
    }
    return NextResponse.json(invoice.toObject());
  } catch (error) {
    console.error("Could not update invoice in MongoDB.", error);
    if (error instanceof SyntaxError) {
      return NextResponse.json({ error: "Invalid JSON request body." }, { status: 400 });
    }
    if (error instanceof Error && error.name === "ValidationError") {
      return NextResponse.json({ error: error.message }, { status: 400 });
    }
    return NextResponse.json({ error: "Could not update invoice." }, { status: 500 });
  }
}

export async function DELETE(_request: Request, { params }: RouteContext) {
  const sessionUser = getSessionUser(_request);
  if (!sessionUser) {
    return NextResponse.json({ error: "Authentication required." }, { status: 401 });
  }
  if (!Types.ObjectId.isValid(params.id)) {
    return NextResponse.json({ error: "Invalid invoice ID." }, { status: 400 });
  }

  try {
    await connectToDatabase();
    const filter =
      sessionUser.role === "admin"
        ? { _id: params.id }
        : { _id: params.id, createdBy: sessionUser.email };
    const invoice = await Invoice.findOneAndDelete(filter);
    if (!invoice) {
      return NextResponse.json({ error: "Invoice not found." }, { status: 404 });
    }
    return NextResponse.json({ success: true });
  } catch (error) {
    console.error("Could not delete invoice from MongoDB.", error);
    return NextResponse.json({ error: "Could not delete invoice." }, { status: 500 });
  }
}
