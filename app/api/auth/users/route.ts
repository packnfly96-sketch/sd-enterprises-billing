import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import connectToDatabase from "@/lib/mongodb";
import User from "@/models/User";

export async function GET(request: Request) {
  const sessionUser = getSessionUser(request);
  if (sessionUser?.role !== "admin") {
    return NextResponse.json({ error: "Admin access required." }, { status: 403 });
  }

  try {
    await connectToDatabase();
    const users = await User.find()
      .select("name email role createdAt")
      .sort({ createdAt: 1 })
      .lean();
    return NextResponse.json(users);
  } catch (error) {
    console.error("Could not fetch team users.", error);
    return NextResponse.json({ error: "Could not fetch users." }, { status: 500 });
  }
}
