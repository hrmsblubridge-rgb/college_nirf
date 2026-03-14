import "@/App.css";
import { BrowserRouter, Routes, Route } from "react-router-dom";
import RegisterPremium from "@/components/RegisterPremium";
import BluBridgeHome from "@/components/BluBridgeHome";

function App() {
  return (
    <div className="App">
      <BrowserRouter>
        <Routes>
          <Route path="/" element={<BluBridgeHome />} />
          <Route path="/register-premium" element={<RegisterPremium />} />
        </Routes>
      </BrowserRouter>
    </div>
  );
}

export default App;
