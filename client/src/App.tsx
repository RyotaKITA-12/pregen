import React from 'react';
import logo from './logo.svg';
import './App.css';
import Button from '@mui/material/Button';

function App() {
  return (
    <div className="App" style={{marginTop: "100px"}}>
      <header className="App-header">
        <p>
          あなたのパワーポイントをアップロードしていただくと、<br></br>
          それに合ったスライドをお作りします。
        </p>
        <Button variant="contained">SUBMIT</Button>
      </header>
    </div>
  );
}

export default App;














