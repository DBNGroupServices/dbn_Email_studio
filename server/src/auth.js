// Verifies the Microsoft Entra token the add-in gets from Outlook single sign-on
// (Office.auth.getAccessToken). Multi-tenant: any work or school account may sign in;
// the token must be issued by Microsoft for OUR app registration.
import { createRemoteJWKSet, jwtVerify } from 'jose';

const JWKS = createRemoteJWKSet(new URL('https://login.microsoftonline.com/common/discovery/v2.0/keys'));

export class AuthError extends Error {}

export function makeVerifier(env) {
  const clientId = env.MS_CLIENT_ID;
  const host = new URL(env.PUBLIC_URL).host;
  const audiences = [clientId, `api://${host}/${clientId}`];
  return async function verify(request) {
    const header = request.headers.get('authorization') || '';
    const token = /^Bearer (.+)$/i.exec(header)?.[1];
    if (!token) throw new AuthError('Sign-in required.');
    let payload;
    try {
      ({ payload } = await jwtVerify(token, JWKS, { audience: audiences, clockTolerance: 60 }));
    } catch {
      throw new AuthError('Your sign-in has expired. Close and reopen the Assistant.');
    }
    const tid = payload.tid, oid = payload.oid;
    const issuers = [`https://login.microsoftonline.com/${tid}/v2.0`, `https://sts.windows.net/${tid}/`];
    if (!tid || !oid || !issuers.includes(payload.iss)) throw new AuthError('This sign-in was not issued for Bulwark HQ Assistant.');
    return {
      id: `${tid}:${oid}`, tid,
      email: String(payload.preferred_username || payload.upn || payload.email || '').slice(0, 320),
      name: String(payload.name || '').slice(0, 200),
    };
  };
}
