import React, { useEffect, useState } from "react";
import { supabase } from "./supabaseClient.js";

const COLORS = {
  paper: "#EDE6D6",
  ink: "#2B2925",
  inkSoft: "#6B6357",
  seal: "#AE3A2A",
  gold: "#A9822F",
  grid: "#C9BC9E",
  card: "#F6F1E4",
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

export function AuthScreen() {
  const [mode, setMode] = useState("login"); // login | signup
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
          {mode === "login" ? "Log in to your account" : "Create an account"}
        </div>

        <label style={{ fontSize: 12, color: COLORS.inkSoft }}>Email</label>
        <input
          type="email"
          required
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          style={inputStyle}
        />

        <label style={{ fontSize: 12, color: COLORS.inkSoft }}>Password</label>
        <input
          type="password"
          required
          minLength={6}
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          style={inputStyle}
        />

        {message && (
          <div
            style={{
              fontSize: 12.5,
              marginTop: 8,
              color: message.type === "error" ? COLORS.seal : "#4C6E52",
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
            color: "#F6EEE2",
            border: "none",
            borderRadius: 7,
            padding: "10px 0",
            fontSize: 14,
            fontWeight: 600,
            cursor: loading ? "default" : "pointer",
            opacity: loading ? 0.6 : 1,
          }}
        >
          {loading ? "Please wait…" : mode === "login" ? "Log in" : "Sign up"}
        </button>

        <div style={{ textAlign: "center", marginTop: 14, fontSize: 12.5 }}>
          {mode === "login" ? (
            <>
              No account?{" "}
              <a href="#" onClick={(e) => { e.preventDefault(); setMode("signup"); setMessage(null); }}>
                Sign up
              </a>
            </>
          ) : (
            <>
              Already have an account?{" "}
              <a href="#" onClick={(e) => { e.preventDefault(); setMode("login"); setMessage(null); }}>
                Log in
              </a>
            </>
          )}
        </div>
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
