import { Types } from "mongoose";
import { NextResponse } from "next/server";
import connectToDatabase from "@/lib/mongodb";
import Customer from "@/models/Customer";

type RouteContext = { params: { id: string } };

export async function DELETE(_request: Request, { params }: RouteContext) {
  if (!Types.ObjectId.isValid(params.id)) {
    return NextResponse.json({ error: "Invalid customer ID." }, { status: 400 });
  }

  try {
    await connectToDatabase();
    const customer = await Customer.findByIdAndDelete(params.id);
    if (!customer) {
      return NextResponse.json({ error: "Customer not found." }, { status: 404 });
    }
    return NextResponse.json({ success: true });
  } catch (error) {
    console.error("Could not delete customer from MongoDB.", error);
    return NextResponse.json({ error: "Could not delete customer." }, { status: 500 });
  }
}
