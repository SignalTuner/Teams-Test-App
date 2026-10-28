import { StrictMode } from "react";
import { createRoot } from "react-dom/client";

import App from "./App";
import { TeamsHostProvider } from "./TeamsHostProvider";

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <TeamsHostProvider><App /></TeamsHostProvider>
  </StrictMode>
);
