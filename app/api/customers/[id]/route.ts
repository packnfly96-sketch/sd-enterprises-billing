import { Types } from "mongoose";
import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import connectToDatabase from "@/lib/mongodb";
import Customer from "@/models/Customer";

type RouteContext = { params: { id: string } };

export async function DELETE(_request: Request, { params }: RouteContext) {
  const sessionUser = getSessionUser(_request);
  if (!sessionUser) {
    return NextResponse.json({ error: "Authentication required." }, { status: 401 });
  }
  if (!Types.ObjectId.isValid(params.id)) {
    return NextResponse.json({ error: "Invalid customer ID." }, { status: 400 });
  }

  try {
    await connectToDatabase();
    const filter =
      sessionUser.role === "admin"
        ? { _id: params.id }
        : { _id: params.id, createdBy: sessionUser.email };
    const customer = await Customer.findOneAndDelete(filter);
    if (!customer) {
      return NextResponse.json({ error: "Customer not found." }, { status: 404 });
    }
    return NextResponse.json({ success: true });
  } catch (error) {
    console.error("Could not delete customer from MongoDB.", error);
    return NextResponse.json({ error: "Could not delete customer." }, { status: 500 });
  }
}
