import { timingSafeEqual } from "node:crypto";
import { hash } from "bcryptjs";
import { NextResponse } from "next/server";
import { getSessionUser } from "@/lib/auth";
import connectToDatabase from "@/lib/mongodb";
import User from "@/models/User";

function isRegistrationAuthorized(request: Request) {
  const expectedSecret = process.env.USER_REGISTRATION_SECRET;
  const providedSecret = request.headers.get("x-registration-secret");
  if (!expectedSecret || !providedSecret) return false;

  const expected = Buffer.from(expectedSecret);
  const provided = Buffer.from(providedSecret);
  return expected.length === provided.length && timingSafeEqual(expected, provided);
}

export async function POST(request: Request) {
  const sessionUser = getSessionUser(request);
  if (sessionUser?.role !== "admin" && !process.env.USER_REGISTRATION_SECRET) {
    return NextResponse.json(
      { error: "User registration is disabled. Configure USER_REGISTRATION_SECRET to enable it." },
      { status: 503 },
    );
  }
  if (sessionUser?.role !== "admin" && !isRegistrationAuthorized(request)) {
    return NextResponse.json({ error: "Registration is not authorized." }, { status: 401 });
  }

  let payload: unknown;
  try {
    payload = await request.json();
  } catch {
    return NextResponse.json({ error: "Invalid JSON request body." }, { status: 400 });
  }

  if (typeof payload !== "object" || payload === null) {
    return NextResponse.json({ error: "Invalid registration data." }, { status: 400 });
  }

  const { name, email, password, role } = payload as Record<string, unknown>;
  if (
    typeof name !== "string" ||
    !name.trim() ||
    typeof email !== "string" ||
    !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim()) ||
    typeof password !== "string" ||
    password.length < 8 ||
    (role !== "admin" && role !== "sub-user")
  ) {
    return NextResponse.json(
      { error: "Provide a name, valid email, password of at least 8 characters, and valid role." },
      { status: 400 },
    );
  }

  try {
    await connectToDatabase();
    const normalizedEmail = email.trim().toLowerCase();
    const existingUser = await User.exists({ email: normalizedEmail });
    if (existingUser) {
      return NextResponse.json({ error: "A user with this email already exists." }, { status: 409 });
    }

    const user = await User.create({
      name: name.trim(),
      email: normalizedEmail,
      password: await hash(password, 12),
      role,
    });

    return NextResponse.json(
      { user: { _id: user._id, name: user.name, email: user.email, role: user.role } },
      { status: 201 },
    );
  } catch (error) {
    console.error("Could not register user in MongoDB.", error);
    if (
      typeof error === "object" &&
      error !== null &&
      "code" in error &&
      error.code === 11000
    ) {
      return NextResponse.json({ error: "A user with this email already exists." }, { status: 409 });
    }
    if (error instanceof Error && error.name === "ValidationError") {
      return NextResponse.json({ error: error.message }, { status: 400 });
    }
    return NextResponse.json({ error: "Could not register user." }, { status: 500 });
  }
}
