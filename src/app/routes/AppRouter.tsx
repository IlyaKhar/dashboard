import { lazy, Suspense, useEffect } from "react";
import { Routes, Route, Navigate } from "react-router-dom";
import Layout from "./Layout";
import { useAuthStore } from "@/store/auth";
import { Loader } from "@/shared/ui/loader";
import { ErrorBoundary } from "@/shared/ui/errorBoundary";

const LoginPage = lazy(() =>
  import("@/pages/login/LoginPage").then((m) => ({ default: m.LoginPage }))
);
const HomePage = lazy(() =>
  import("@/pages/home/HomePage").then((m) => ({ default: m.HomePage }))
);
const DepartmentsPage = lazy(() =>
  import("@/pages/departments/DepartmentsPage").then((m) => ({
    default: m.DepartmentsPage,
  }))
);
const GroupsPage = lazy(() =>
  import("@/pages/groups/GroupsPage").then((m) => ({ default: m.GroupsPage }))
);
const StoryPage = lazy(() =>
  import("@/pages/story/StoryPage").then((m) => ({ default: m.StoryPage }))
);
const ByLessonPage = lazy(() =>
  import("@/pages/byLesson/ByLessonPage").then((m) => ({
    default: m.ByLessonPage,
  }))
);
const NotFound = lazy(() =>
  import("./NotFound").then((m) => ({ default: m.NotFound }))
);

function Logout() {
  const logout = useAuthStore((state) => state.logout);

  useEffect(() => {
    logout();
  }, [logout]);

  return <Navigate to="/login" replace />;
}

function PageSuspense({ children }: { children: React.ReactNode }) {
  return (
    <ErrorBoundary>
      <Suspense fallback={<Loader text="Загрузка страницы..." />}>
        {children}
      </Suspense>
    </ErrorBoundary>
  );
}

export function AppRouter() {
  const isAuth = useAuthStore((state) => state.isAuth);

  return (
    <Routes>
      <Route
        path="/login"
        element={
          <PageSuspense>
            <LoginPage />
          </PageSuspense>
        }
      />

      {isAuth ? (
        <>
          <Route path="/logout" element={<Logout />} />
          <Route element={<Layout />}>
            <Route
              index
              element={
                <PageSuspense>
                  <HomePage />
                </PageSuspense>
              }
            />
            <Route
              path="/departments"
              element={
                <PageSuspense>
                  <DepartmentsPage />
                </PageSuspense>
              }
            />
            <Route
              path="/groups"
              element={
                <PageSuspense>
                  <GroupsPage />
                </PageSuspense>
              }
            />
            <Route
              path="/story"
              element={
                <PageSuspense>
                  <StoryPage />
                </PageSuspense>
              }
            />
            <Route
              path="/by-lesson"
              element={
                <PageSuspense>
                  <ByLessonPage />
                </PageSuspense>
              }
            />
          </Route>
          <Route
            path="*"
            element={
              <PageSuspense>
                <NotFound />
              </PageSuspense>
            }
          />
        </>
      ) : (
        <Route path="*" element={<Navigate to="/login" replace />} />
      )}
    </Routes>
  );
}
