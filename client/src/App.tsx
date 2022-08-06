import React, { useState } from 'react';
import logo from './logo.svg';
import './App.css';
import Button from '@mui/material/Button';
import { SendButton } from './SendButton';

function App() {
    const [postFileData, setPostFileData] = useState({})
    const [filename,setFilename] = useState("")
    const changeUploadFile = async (event:any) => {
      const { name, files } = event.target;
      setPostFileData({
        ...postFileData,
        [name]: files[0],
      });
      setFilename(
        files[0].name
      )

      event.target.value = '';
    };

  return (
    <div className="App" style={{marginTop: "100px"}}>
       <header className="App-header">
        <p>
          あなたのパワーポイントをアップロードしていただくと、<br></br>
          それに合ったスライドをお作りします。
        </p>
        <SendButton onChange={changeUploadFile} name="pptx">GO!</SendButton>
        {filename}
        <Button></Button>
      </header>
    </div>
  );
}

export default App;
















