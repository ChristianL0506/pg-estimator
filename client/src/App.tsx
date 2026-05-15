import { useState, useCallback } from "react";
import { Switch, Route, Router } from "wouter";
import { useHashLocation } from "wouter/use-hash-location";
import { queryClient, setAuthToken } from "./lib/queryClient";
import { QueryClientProvider } from "@tanstack/react-query";
import { Toaster } from "@/components/ui/toaster";
import { TooltipProvider } from "@/components/ui/tooltip";
import NotFound from "@/pages/not-found";
import DashboardPage from "@/pages/DashboardPage";
import TakeoffPage from "@/pages/TakeoffPage";
import EstimatingPage from "@/pages/EstimatingPage";
import LoginPage from "@/pages/LoginPage";
import SettingsPage from "@/pages/SettingsPage";
import CostDatabasePage from "@/pages/CostDatabasePage";
import ProjectHistoryPage from "@/pages/ProjectHistoryPage";
import BidDashboardPage from "@/pages/BidDashboardPage";
import HelpPage from "@/pages/HelpPage";
import MethodsPage from "@/pages/MethodsPage";
import PerformancePage from "@/pages/PerformancePage";

function AppRouter() {
  return (
    <Switch>
      <Route path="/" component={DashboardPage} />
      <Route path="/mechanical">
        {() => <TakeoffPage discipline="mechanical" />}
      </Route>
      <Route path="/structural">
        {() => <TakeoffPage discipline="structural" />}
      </Route>
      <Route path="/civil">
        {() => <TakeoffPage discipline="civil" />}
      </Route>
      <Route path="/estimating" component={EstimatingPage} />
      <Route path="/methods" component={MethodsPage} />
      <Route path="/cost-database" component={CostDatabasePage} />
      <Route path="/project-history" component={ProjectHistoryPage} />
      <Route path="/bids" component={BidDashboardPage} />
      <Route path="/performance" component={PerformancePage} />
      <Route path="/help" component={HelpPage} />
      <Route path="/settings" component={SettingsPage} />
      <Route component={NotFound} />
    </Switch>
  );
}

function App() {
  const [authToken, setAuthTokenState] = useState<string | null>(null);
  const [_username, setUsername] = useState<string | null>(null);

  const handleLogin = useCallback((token: string, username: string) => {
    setAuthTokenState(token);
    setAuthToken(token);
    setUsername(username);
    (window as any).__PG_AUTH_TOKEN__ = token;
    queryClient.clear();
  }, []);

  const handleLogout = useCallback(async () => {
    try {
      const API_BASE = "__PORT_5000__".startsWith("__") ? "" : "__PORT_5000__";
      await fetch(`${API_BASE}/api/logout`, {
        method: "POST",
        headers: { Authorization: `Bearer ${authToken}` },
      });
    } catch {}
    setAuthTokenState(null);
    setAuthToken(null);
    setUsername(null);
    queryClient.clear();
  }, [authToken]);

  if (!authToken) {
    return <LoginPage onLogin={handleLogin} />;
  }

  return (
    <QueryClientProvider client={queryClient}>
      <TooltipProvider>
        <Toaster />
        <Router hook={useHashLocation}>
          <AppRouter />
        </Router>
      </TooltipProvider>
    </QueryClientProvider>
  );
}

export default App;
