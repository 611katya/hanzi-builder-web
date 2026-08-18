import React, { useState } from "react";
import { useSession, AuthScreen } from "./Auth.jsx";
import { supabase } from "./supabaseClient.js";
import HanziBuilder from "./HanziBuilder.jsx";

export default function App() {
  const session = useSession();
  const [showAuth, setShowAuth] = useState(false);

  if (session === undefined) {
    return <div style={{ padding: 40, textAlign: "center", fontFamily: "system-ui" }}>Loading…</div>;
  }

  // Logged in: always show the full app, regardless of how they got here.
  if (session) {
    return (
      <div>
        <TopBar>
          <span>{session.user.email}</span>
          <button onClick={() => supabase.auth.signOut()} style={buttonStyle}>
            Log out
          </button>
        </TopBar>
        <HanziBuilder userId={session.user.id} />
      </div>
    );
  }

  // Not logged in, and the person chose to log in / sign up right now.
  if (showAuth) {
    return <AuthScreen onGuest={() => setShowAuth(false)} />;
  }

  // Default landing experience: play immediately, no account required.
  return (
    <div>
      <TopBar>
        <span>Browsing as guest — changes won't be saved</span>
        <button onClick={() => setShowAuth(true)} style={{ ...buttonStyle, background: "#556B2F", borderColor: "#556B2F", color: "#FBF9EF" }}>
          Log in / Sign up
        </button>
      </TopBar>
      <HanziBuilder userId={null} />
    </div>
  );
}

function TopBar({ children }) {
  return (
    <div
      style={{
        display: "flex",
        justifyContent: "flex-end",
        alignItems: "center",
        gap: 10,
        padding: "8px 16px",
        fontSize: 12.5,
        fontFamily: "system-ui, sans-serif",
        color: "#7A7A56",
      }}
    >
      {children}
    </div>
  );
}

const buttonStyle = {
  border: "1px solid #DDD9BB",
  background: "transparent",
  borderRadius: 6,
  padding: "4px 10px",
  fontSize: 12,
  cursor: "pointer",
};
