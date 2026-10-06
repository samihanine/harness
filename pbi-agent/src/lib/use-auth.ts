import { useCallback, useEffect, useState } from "react";
import type { AuthStatus } from "@/pbi/auth";
import { getStatus } from "@/pbi/auth";

/** Power BI sign-in state (polled while a sign-in is pending). */
export function useAuth() {
  const [status, setStatus] = useState<AuthStatus | null>(null);
  const [error, setError] = useState("");
  const refresh = useCallback(
    () =>
      getStatus()
        .then((s) => {
          setStatus(s);
          setError("");
        })
        .catch((e: Error) => setError(e.message)),
    [],
  );
  useEffect(() => {
    void refresh();
  }, [refresh]);
  useEffect(() => {
    if (!status?.pending) return;
    const timer = setInterval(() => void refresh(), 3000);
    return () => clearInterval(timer);
  }, [status?.pending, refresh]);
  return { status, error, refresh };
}
