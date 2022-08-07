import React, { useState } from "react";
import logo from "./logo.svg";
import logo1 from "./download 2.jpg";
import logo2 from "./logo.png";
import "./App.css";
import Button from "@mui/material/Button";
import { SendButton } from "./SendButton";
import axios from "./utils/http_client";
import Link from "@mui/material/Link";
import { Download } from "@mui/icons-material";

function App() {
  const [postFileData, setPostFileData] = useState<any>({});
  const [filename, setFilename] = useState("");
  const changeUploadFile = async (event: any) => {
    const { name, files } = event.target;
    setPostFileData({
      ...postFileData,
      [name]: files[0],
    });
    setFilename(files[0].name);

    event.target.value = "";
  };

  const sendFile = async () => {
    const formData = new FormData();
    formData.append("file", postFileData.pptx ? postFileData.pptx : "");
    const response = await axios
      .post("/files/upload_pptx", formData, {
        headers: {
          "Content-Type": "multipart/form-data",
        },
        responseType: "blob",
      })
      .then((res) => {
        const url = window.URL.createObjectURL(new Blob([res.data]));
        const link = document.createElement("a");
        link.href = url;
        link.setAttribute("download", "pregen.pptx");
        document.body.appendChild(link);
        link.click();
      });
    console.log(response);
  };

  const download = async () => {
    await axios
      .get("/static/template.pptx", {
        responseType: "blob",
      })
      .then((res) => {
        const url = window.URL.createObjectURL(new Blob([res.data]));
        const link = document.createElement("a");
        link.href = url;
        link.setAttribute("download", "template.pptx");
        document.body.appendChild(link);
        link.click();
      });
  };

  return (
    <div className="App" style={{ marginTop: "100px" }}>
      <header className="App-header">
        <img className="logo" src={logo2} alt="picture" />
        <p>
          <Link underline="hover" onClick={download}>
            こちらのテンプレート
          </Link>
          を編集して<br></br>
          アップロードしていただくと、<br></br>
          それに合ったスライドを作成します。
        </p>

        <div className="Button">
          <SendButton onChange={changeUploadFile} name="pptx">
            アップロード
          </SendButton>
          <Button onClick={sendFile} variant="contained" disabled={!filename}>
            送信
          </Button>
          <div>{filename}</div>
          <img src={logo1} alt="picture" />
        </div>
      </header>
    </div>
  );
}

export default App;
