export type UserRole = "admin" | "user";

export interface AuthStore {
  isAuth: boolean;
  token: string | null;
  role: UserRole | null;
  login: (token: string, role: UserRole) => void;
  logout: () => void;
}
