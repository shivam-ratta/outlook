import "./App.css";
import OutlookMail from "./components/OutlookMail";
import UserState from "./context/UserState";

function App() {
  return (
    <div className="App">
      <UserState>
        <OutlookMail />
      </UserState>
    </div>
  );
}

export default App;
