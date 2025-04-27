import React from "react";

const STORAGE_KEY = "orientation";

export const App: React.FC = () => {
  const [mode, setMode] = React.useState<string>(
    localStorage.getItem(STORAGE_KEY) ?? "row"
  );

  const update = (value: string) => {
    setMode(value);
    localStorage.setItem(STORAGE_KEY, value);
    Office.context.document.settings.refreshAsync(); // trigger recalc
  };

  return (
    <div className="ms-welcome">
      <h2 className="mt-0">Factorial orientation</h2>

      <label>
        <input
          type="radio"
          checked={mode === "row"}
          onChange={() => update("row")}
        />{" "}
        Row (→)
      </label>
      <br />
      <label>
        <input
          type="radio"
          checked={mode === "column"}
          onChange={() => update("column")}
        />{" "}
        Column (↓)
      </label>

      <p className="mt-18">
        Try <code>=TESTVELIXO.FACTORIALROW(7)</code> and switch the radio — Excel
        will respill automatically.
      </p>
    </div>
  );
};
