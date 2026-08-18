import React, { useState } from "react";
import { useSession, AuthScreen } from "./Auth.jsx";
import { supabase } from "./supabaseClient.js";
import HanziBuilder from "./HanziBuilder.jsx";

export default function App() {
  const session = useSession();
  const [guest, setGuest] = useState(false);

  if (session === undefined) {
    return <div style={{ padding: 40, textAlign: "center", fontFamily: "system-ui" }}>Loading…</div>;
  }

  if (!session && !guest) {
    return <AuthScreen onGuest={() => setGuest(true)} />;
  }

  const isGuest = !session && guest;

  return (
    <div>
      <div
        style={{
          display: "flex",
          justifyContent: "flex-end",
          alignItems: "center",
          gap: 10,
          padding: "8px 16px",
          fontSize: 12.5,
          fontFamily: "system-ui, sans-serif",
          color: "#6B6357",
        }}
      >
        {isGuest ? (
          <>
            <span>Browsing as guest — changes won't be saved</span>
            <button
              onClick={() => setGuest(false)}
              style={{
                border: "1px solid #AE3A2A",
                background: "#AE3A2A",
                color: "#F6EEE2",
                borderRadius: 6,
                padding: "4px 10px",
                fontSize: 12,
                cursor: "pointer",
              }}
            >
              Sign up to save progress
            </button>
          </>
        ) : (
          <>
            <span>{session.user.email}</span>
            <button
              onClick={() => supabase.auth.signOut()}
              style={{
                border: "1px solid #C9BC9E",
                background: "transparent",
                borderRadius: 6,
                padding: "4px 10px",
                fontSize: 12,
                cursor: "pointer",
              }}
            >
              Log out
            </button>
          </>
        )}
      </div>
      <HanziBuilder userId={isGuest ? null : session.user.id} />
    </div>
  );
}
