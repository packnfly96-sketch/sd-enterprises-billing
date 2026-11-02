import { NextResponse } from "next/server";
import connectToDatabase from "@/lib/mongodb";
import Customer from "@/models/Customer";

export async function GET() {
  try {
    await connectToDatabase();
    const customers = await Customer.find().sort({ createdAt: -1 }).lean();
    return NextResponse.json(customers);
  } catch (error) {
    console.error("Could not fetch customers from MongoDB.", error);
    return NextResponse.json({ error: "Could not fetch customers." }, { status: 500 });
  }
}

export async function POST(request: Request) {
  try {
    const customerData: unknown = await request.json();
    if (
      typeof customerData !== "object" ||
      customerData === null ||
      typeof (customerData as Record<string, unknown>).name !== "string" ||
      typeof (customerData as Record<string, unknown>).phone !== "string" ||
      typeof (customerData as Record<string, unknown>).address !== "string" ||
      typeof (customerData as Record<string, unknown>).state !== "string"
    ) {
      return NextResponse.json({ error: "Invalid customer data." }, { status: 400 });
    }

    await connectToDatabase();
    const customer = await Customer.create(customerData);
    return NextResponse.json(customer.toObject(), { status: 201 });
  } catch (error) {
    console.error("Could not create customer in MongoDB.", error);
    if (error instanceof SyntaxError) {
      return NextResponse.json({ error: "Invalid JSON request body." }, { status: 400 });
    }
    if (error instanceof Error && error.name === "ValidationError") {
      return NextResponse.json({ error: error.message }, { status: 400 });
    }
    return NextResponse.json({ error: "Could not create customer." }, { status: 500 });
  }
}
