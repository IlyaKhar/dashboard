import type { AuthStore, UserRole } from "./types";
import { create } from "zustand";
import { persist } from "zustand/middleware";

export const useAuthStore = create<AuthStore>()(
  persist(
    (set) => ({
      isAuth: false,
      token: null,
      role: null,

      login: (token: string, role: UserRole) =>
        set({
          token,
          role,
          isAuth: true,
        }),

      logout: () =>
        set({
          token: null,
          role: null,
          isAuth: false,
        }),
    }),
    {
      name: "auth-store",
    }
  )
);
