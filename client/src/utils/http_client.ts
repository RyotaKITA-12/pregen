import { Axios } from "axios";

const axios = new Axios({
  baseURL: process.env.REACT_APP_API_BASE_URL || "http://localhost:8080",
});

export default axios;
