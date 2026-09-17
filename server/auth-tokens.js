import { SignJWT, jwtVerify } from 'jose';
import { getCookie, setCookie, deleteCookie } from 'hono/cookie';

const COOKIE_NAME = 'bicubo_token';

function getSecret() {
  const secret = process.env.JWT_SECRET || 'bicubo-dev-secret-change-me';
  return new TextEncoder().encode(secret);
}

export async function signToken(user) {
  return new SignJWT({
    sub: user.id,
    email: user.email,
    name: user.name,
    role: user.role,
  })
    .setProtectedHeader({ alg: 'HS256' })
    .setIssuedAt()
    .setExpirationTime('7d')
    .sign(getSecret());
}

export async function verifyToken(token) {
  if (!token) return null;
  try {
    const { payload } = await jwtVerify(token, getSecret());
    return {
      id: payload.sub,
      email: payload.email,
      name: payload.name,
      role: payload.role,
    };
  } catch {
    return null;
  }
}

export function readAuthToken(c) {
  const header = c.req.header('authorization') || '';
  if (header.startsWith('Bearer ')) return header.slice(7);
  return getCookie(c, COOKIE_NAME) || null;
}

export function setAuthCookie(c, token) {
  setCookie(c, COOKIE_NAME, token, {
    httpOnly: true,
    secure: process.env.VERCEL === '1' || process.env.NODE_ENV === 'production',
    sameSite: 'Lax',
    path: '/',
    maxAge: 60 * 60 * 24 * 7,
  });
}

export function clearAuthCookie(c) {
  deleteCookie(c, COOKIE_NAME, { path: '/' });
}

export { COOKIE_NAME };
