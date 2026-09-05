import { getIronSession } from "iron-session";
import { cookies } from "next/headers";

export interface SessionData {
  isLoggedIn: boolean;
  branchId: string;
  compId: string;
  userId: string;
  finyear: string;
}

export const defaultSession: SessionData = {
  isLoggedIn: false,
  branchId: "",
  compId: "",
  userId: "",
  finyear: "",
};

export const sessionOptions = {
  password: process.env.SECRET_COOKIE_PASSWORD || "complex_password_at_least_32_characters_long",
  cookieName: "erp_v3_session",
  cookieOptions: {
    secure: process.env.NODE_ENV === "production",
  },
};

export async function getSession() {
  const cookieStore = await cookies();
  const session = await getIronSession<SessionData>(cookieStore, sessionOptions);

  if (!session.isLoggedIn) {
    session.isLoggedIn = defaultSession.isLoggedIn;
    session.branchId = defaultSession.branchId;
    session.compId = defaultSession.compId;
  }

  return session;
}
