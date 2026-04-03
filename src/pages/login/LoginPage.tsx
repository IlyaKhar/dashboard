import { useState } from "react";
import type { FormEvent, ChangeEvent } from "react";
import { Button } from "@/shared/ui/button";
import { Input } from "@/shared/ui/input";
import { Label } from "@/shared/ui/label";
import { Loader } from "@/shared/ui/loader";
import { Eye, EyeOff, LogIn } from "lucide-react";
import { useAuthStore } from "@/store/auth";
import { useNavigate } from "react-router-dom";
import { loginRequest } from "@/shared/api";

export function LoginPage() {
  const [showPassword, setShowPassword] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const navigate = useNavigate();
  const login = useAuthStore((state) => state.login);

  const [formData, setFormData] = useState({
    username: "",
    password: "",
  });

  const handleSubmit = async (e: FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    setIsLoading(true);
    setErrorMessage(null);

    try {
      const response = await loginRequest({
        username: formData.username,
        password: formData.password,
      });

      login(response.token, response.role as "admin" | "user");
      navigate("/");
    } catch (err) {
      const message =
        err instanceof Error ? err.message : "Не удалось войти в систему";
      setErrorMessage(message);
    } finally {
      setIsLoading(false);
    }
  };

  const handleChange = (e: ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target;
    setFormData((prev) => ({
      ...prev,
      [name]: value,
    }));
    if (errorMessage) setErrorMessage(null);
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-linear-to-br from-gray-50 to-gray-100 p-4">
      <div className="w-full max-w-md rounded-2xl border border-gray-300 bg-white shadow-xl p-6 flex flex-col gap-6">
        <div className="flex flex-col items-center gap-4 text-center">
          <img src="/images/logo.png" alt="Logo" className="h-16 w-16" />
          <div>
            <h1 className="text-2xl font-bold text-black">Вход в систему</h1>
            <p className="text-gray-600 text-sm">
              Введите свои учетные данные для доступа
            </p>
          </div>
        </div>

        {errorMessage && (
          <div className="rounded-lg border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-600">
            {errorMessage}
          </div>
        )}

        <form className="flex flex-col gap-5" onSubmit={handleSubmit}>
          <div className="flex flex-col gap-2">
            <Label htmlFor="username" className="text-black">
              Имя пользователя:
            </Label>
            <Input
              id="username"
              name="username"
              type="text"
              placeholder="Введите логин"
              value={formData.username}
              onChange={handleChange}
              required
              disabled={isLoading}
              className="border-gray-300!"
            />
          </div>

          <div className="flex flex-col gap-2">
            <Label htmlFor="password" className="text-black">
              Пароль:
            </Label>
            <div className="relative">
              <Input
                id="password"
                name="password"
                type={showPassword ? "text" : "password"}
                placeholder="Введите пароль"
                value={formData.password}
                onChange={handleChange}
                required
                disabled={isLoading}
                className="pr-10 border-gray-300!"
              />
              <button
                type="button"
                onClick={() => setShowPassword(!showPassword)}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-500 hover:text-black"
                aria-label={showPassword ? "Скрыть пароль" : "Показать пароль"}
              >
                {showPassword ? (
                  <EyeOff className="h-4 w-4" />
                ) : (
                  <Eye className="h-4 w-4" />
                )}
              </button>
            </div>
          </div>

          <Button
            type="submit"
            disabled={isLoading}
            className="w-full bg-black hover:bg-gray-800 text-white h-10"
          >
            {isLoading ? (
              <Loader className="py-0" text="Вход..." />
            ) : (
              <span className="flex items-center justify-center">
                <LogIn className="mr-2 h-4 w-4" />
                Войти
              </span>
            )}
          </Button>
        </form>
      </div>
    </div>
  );
}
