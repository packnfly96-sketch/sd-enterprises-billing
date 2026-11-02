import { compare, hash } from "bcryptjs";
import { NextResponse } from "next/server";
import {
  createSessionToken,
  getSessionCookieOptions,
  SESSION_COOKIE_NAME,
} from "@/lib/auth";
import connectToDatabase from "@/lib/mongodb";
import User from "@/models/User";

export async function POST(request: Request) {
  let credentials: unknown;

  try {
    credentials = await request.json();
  } catch {
    return NextResponse.json({ error: "Invalid JSON request body." }, { status: 400 });
  }

  if (
    typeof credentials !== "object" ||
    credentials === null ||
    typeof (credentials as Record<string, unknown>).email !== "string" ||
    typeof (credentials as Record<string, unknown>).password !== "string"
  ) {
    return NextResponse.json({ error: "Email and password are required." }, { status: 400 });
  }

  const { email, password } = credentials as { email: string; password: string };
  const normalizedEmail = email.trim().toLowerCase();
  if (!normalizedEmail || !password) {
    return NextResponse.json({ error: "Email and password are required." }, { status: 400 });
  }

  try {
    await connectToDatabase();
    const user = await User.findOne({ email: normalizedEmail }).select("+password");
    if (!user) {
      return NextResponse.json({ error: "Invalid email or password." }, { status: 401 });
    }

    const isHashedPassword = /^\$2[aby]\$\d{2}\$/.test(user.password);
    const passwordMatches = isHashedPassword
      ? await compare(password, user.password)
      : password === user.password;

    if (!passwordMatches) {
      return NextResponse.json({ error: "Invalid email or password." }, { status: 401 });
    }

    if (!isHashedPassword) {
      user.password = await hash(password, 12);
      await user.save();
    }

    const response = NextResponse.json({
      user: {
        name: user.name,
        email: user.email,
        role: user.role,
      },
    });
    response.cookies.set(
      SESSION_COOKIE_NAME,
      createSessionToken({
        name: user.name,
        email: user.email,
        role: user.role,
      }),
      getSessionCookieOptions(),
    );
    return response;
  } catch (error) {
    console.error("Could not authenticate user against MongoDB.", error);
    return NextResponse.json({ error: "Login is temporarily unavailable." }, { status: 500 });
  }
}
