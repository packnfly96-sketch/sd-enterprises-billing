import { createHmac, timingSafeEqual } from "node:crypto";

export const SESSION_COOKIE_NAME = "agc_session";
const SESSION_MAX_AGE_SECONDS = 60 * 60 * 24 * 7;

export type SessionUser = {
  name: string;
  email: string;
  role: "admin" | "sub-user";
};

function getSessionSecret() {
  const secret = process.env.AUTH_SECRET ?? process.env.USER_REGISTRATION_SECRET;
  if (!secret) {
    throw new Error("Set AUTH_SECRET to enable authenticated sessions.");
  }
  return secret;
}

function sign(payload: string) {
  return createHmac("sha256", getSessionSecret()).update(payload).digest("base64url");
}

export function createSessionToken(user: SessionUser) {
  const payload = Buffer.from(
    JSON.stringify({ ...user, expiresAt: Date.now() + SESSION_MAX_AGE_SECONDS * 1000 }),
  ).toString("base64url");
  return `${payload}.${sign(payload)}`;
}

export function getSessionCookieOptions() {
  return {
    httpOnly: true,
    secure: process.env.NODE_ENV === "production",
    sameSite: "lax" as const,
    path: "/",
    maxAge: SESSION_MAX_AGE_SECONDS,
  };
}

export function clearSessionCookieOptions() {
  return { ...getSessionCookieOptions(), maxAge: 0 };
}

export function getSessionUser(request: Request): SessionUser | null {
  const cookie = request.headers
    .get("cookie")
    ?.split(";")
    .map((part) => part.trim())
    .find((part) => part.startsWith(`${SESSION_COOKIE_NAME}=`));

  if (!cookie) return null;

  try {
    const token = decodeURIComponent(cookie.slice(SESSION_COOKIE_NAME.length + 1));
    const separator = token.lastIndexOf(".");
    if (separator <= 0) return null;

    const payload = token.slice(0, separator);
    const providedSignature = Buffer.from(token.slice(separator + 1));
    const expectedSignature = Buffer.from(sign(payload));
    if (
      providedSignature.length !== expectedSignature.length ||
      !timingSafeEqual(providedSignature, expectedSignature)
    ) {
      return null;
    }

    const decoded: unknown = JSON.parse(Buffer.from(payload, "base64url").toString("utf8"));
    if (typeof decoded !== "object" || decoded === null) return null;

    const session = decoded as Record<string, unknown>;
    if (
      typeof session.name !== "string" ||
      typeof session.email !== "string" ||
      (session.role !== "admin" && session.role !== "sub-user") ||
      typeof session.expiresAt !== "number" ||
      session.expiresAt <= Date.now()
    ) {
      return null;
    }

    return {
      name: session.name,
      email: session.email,
      role: session.role,
    };
  } catch {
    return null;
  }
}
