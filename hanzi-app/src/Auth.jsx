import React, { useEffect, useState } from "react";
import { supabase } from "./supabaseClient.js";

const COLORS = {
  paper: "#F2F0E1",
  ink: "#2A2A16",
  inkSoft: "#7A7A56",
  seal: "#556B2F",
  gold: "#584C25",
  grid: "#DDD9BB",
  card: "#FAF9EF",
  error: "#A6432E",
};

export function useSession() {
  const [session, setSession] = useState(undefined); // undefined = loading, null = logged out

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: listener } = supabase.auth.onAuthStateChange((_event, newSession) => {
      setSession(newSession);
    });
    return () => listener.subscription.unsubscribe();
  }, []);

  return session;
}

// Detects the moment someone arrives back on the site via a "reset your
// password" email link. Supabase signs them into a temporary session and
// fires this event — the app should show a "set new password" screen
// instead of the normal logged-in view until they've done that.
export function useRecoveryMode() {
  const [recovery, setRecovery] = useState(false);

  useEffect(() => {
    const { data: listener } = supabase.auth.onAuthStateChange((event) => {
      if (event === "PASSWORD_RECOVERY") setRecovery(true);
    });
    return () => listener.subscription.unsubscribe();
  }, []);

  return [recovery, setRecovery];
}

export function AuthScreen({ onGuest }) {
  const [mode, setMode] = useState("login"); // login | signup | forgot
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [message, setMessage] = useState(null);
  const [loading, setLoading] = useState(false);

  async function handleSubmit(e) {
    e.preventDefault();
    setMessage(null);
    setLoading(true);
    try {
      if (mode === "signup") {
        const { error } = await supabase.auth.signUp({ email, password });
        if (error) throw error;
        setMessage({
          type: "success",
          text: "Account created. Check your email to confirm, then log in.",
        });
        setMode("login");
      } else if (mode === "forgot") {
        const { error } = await supabase.auth.resetPasswordForEmail(email, {
          redirectTo: window.location.origin,
        });
        if (error) throw error;
        setMessage({
          type: "success",
          text: "If that email has an account, a reset link has been sent. Check your inbox (and spam).",
        });
      } else {
        const { error } = await supabase.auth.signInWithPassword({ email, password });
        if (error) throw error;
      }
    } catch (err) {
      setMessage({ type: "error", text: err.message || "Something went wrong." });
    } finally {
      setLoading(false);
    }
  }

  return (
    <div
      style={{
        minHeight: "100vh",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        background: COLORS.paper,
        fontFamily: "system-ui, sans-serif",
        padding: 20,
      }}
    >
      <form
        onSubmit={handleSubmit}
        style={{
          background: COLORS.card,
          border: `1px solid ${COLORS.grid}`,
          borderRadius: 10,
          padding: 28,
          width: 320,
        }}
      >
        <div style={{ textAlign: "center", fontSize: 26, marginBottom: 4 }}>拼字游戏</div>
        <div style={{ textAlign: "center", fontSize: 13, color: COLORS.inkSoft, marginBottom: 20 }}>
          {mode === "login" ? "Log in to your account" : mode === "signup" ? "Create an account" : "Reset your password"}
        </div>

        <label style={{ fontSize: 12, color: COLORS.inkSoft }}>Email</label>
        <input
          type="email"
          required
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          style={inputStyle}
        />

        {mode !== "forgot" && (
          <>
            <label style={{ fontSize: 12, color: COLORS.inkSoft }}>Password</label>
            <input
              type="password"
              required
              minLength={6}
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              style={inputStyle}
            />
          </>
        )}

        {mode === "login" && (
          <div style={{ textAlign: "right", marginTop: -6, marginBottom: 12 }}>
            <a
              href="#"
              style={{ fontSize: 11.5, color: COLORS.inkSoft }}
              onClick={(e) => { e.preventDefault(); setMode("forgot"); setMessage(null); }}
            >
              Forgot password?
            </a>
          </div>
        )}

        {message && (
          <div
            style={{
              fontSize: 12.5,
              marginTop: 8,
              color: message.type === "error" ? COLORS.error : "#3F3F00",
              fontWeight: 600,
            }}
          >
            {message.text}
          </div>
        )}

        <button
          type="submit"
          disabled={loading}
          style={{
            width: "100%",
            marginTop: 16,
            background: COLORS.seal,
            color: "#FBF9EF",
            border: "none",
            borderRadius: 7,
            padding: "10px 0",
            fontSize: 14,
            fontWeight: 600,
            cursor: loading ? "default" : "pointer",
            opacity: loading ? 0.6 : 1,
          }}
        >
          {loading ? "Please wait…" : mode === "login" ? "Log in" : mode === "signup" ? "Sign up" : "Send reset link"}
        </button>

        <div style={{ textAlign: "center", marginTop: 14, fontSize: 12.5 }}>
          {mode === "login" ? (
            <>
              No account?{" "}
              <a href="#" onClick={(e) => { e.preventDefault(); setMode("signup"); setMessage(null); }}>
                Sign up
              </a>
            </>
          ) : mode === "signup" ? (
            <>
              Already have an account?{" "}
              <a href="#" onClick={(e) => { e.preventDefault(); setMode("login"); setMessage(null); }}>
                Log in
              </a>
            </>
          ) : (
            <a href="#" onClick={(e) => { e.preventDefault(); setMode("login"); setMessage(null); }}>
              ← Back to log in
            </a>
          )}
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: 10, margin: "18px 0 14px" }}>
          <div style={{ flex: 1, borderTop: `1px solid ${COLORS.grid}` }} />
          <span style={{ fontSize: 11, color: COLORS.inkSoft }}>or</span>
          <div style={{ flex: 1, borderTop: `1px solid ${COLORS.grid}` }} />
        </div>

        <button
          type="button"
          onClick={onGuest}
          style={{
            width: "100%",
            background: "transparent",
            border: `1px solid ${COLORS.grid}`,
            color: COLORS.inkSoft,
            borderRadius: 7,
            padding: "10px 0",
            fontSize: 13,
            cursor: "pointer",
          }}
        >
          ← Continue browsing without an account
        </button>
      </form>
    </div>
  );
}

const inputStyle = {
  display: "block",
  width: "100%",
  boxSizing: "border-box",
  border: `1px solid ${COLORS.grid}`,
  borderRadius: 6,
  padding: "8px 10px",
  fontSize: 14,
  marginTop: 4,
  marginBottom: 12,
};

// Shown instead of the normal app once useRecoveryMode() has fired,
// i.e. the person arrived here by clicking the link in a password-reset email.
export function ResetPasswordScreen({ onDone }) {
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [message, setMessage] = useState(null);
  const [loading, setLoading] = useState(false);

  async function handleSubmit(e) {
    e.preventDefault();
    setMessage(null);
    if (password.length < 6) {
      setMessage({ type: "error", text: "Password must be at least 6 characters." });
      return;
    }
    if (password !== confirm) {
      setMessage({ type: "error", text: "Passwords don't match." });
      return;
    }
    setLoading(true);
    const { error } = await supabase.auth.updateUser({ password });
    setLoading(false);
    if (error) {
      setMessage({ type: "error", text: error.message });
      return;
    }
    setMessage({ type: "success", text: "Password updated — taking you in…" });
    setTimeout(() => onDone(), 1200);
  }

  return (
    <div
      style={{
        minHeight: "100vh",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        background: COLORS.paper,
        fontFamily: "system-ui, sans-serif",
        padding: 20,
      }}
    >
      <form
        onSubmit={handleSubmit}
        style={{
          background: COLORS.card,
          border: `1px solid ${COLORS.grid}`,
          borderRadius: 10,
          padding: 28,
          width: 320,
        }}
      >
        <div style={{ textAlign: "center", fontSize: 26, marginBottom: 4 }}>拼字游戏</div>
        <div style={{ textAlign: "center", fontSize: 13, color: COLORS.inkSoft, marginBottom: 20 }}>
          Choose a new password
        </div>

        <label style={{ fontSize: 12, color: COLORS.inkSoft }}>New password</label>
        <input
          type="password"
          required
          minLength={6}
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          style={inputStyle}
        />

        <label style={{ fontSize: 12, color: COLORS.inkSoft }}>Confirm new password</label>
        <input
          type="password"
          required
          minLength={6}
          value={confirm}
          onChange={(e) => setConfirm(e.target.value)}
          style={inputStyle}
        />

        {message && (
          <div
            style={{
              fontSize: 12.5,
              marginTop: 8,
              color: message.type === "error" ? COLORS.error : "#3F3F00",
              fontWeight: 600,
            }}
          >
            {message.text}
          </div>
        )}

        <button
          type="submit"
          disabled={loading}
          style={{
            width: "100%",
            marginTop: 16,
            background: COLORS.seal,
            color: "#FBF9EF",
            border: "none",
            borderRadius: 7,
            padding: "10px 0",
            fontSize: 14,
            fontWeight: 600,
            cursor: loading ? "default" : "pointer",
            opacity: loading ? 0.6 : 1,
          }}
        >
          {loading ? "Please wait…" : "Update password"}
        </button>
      </form>
    </div>
  );
}
