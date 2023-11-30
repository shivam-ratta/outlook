import * as sweetalert from "sweetalert";
import axios from "axios";

const baseURL = process.env.REACT_APP_API_URL;

const accessToken = () => {return localStorage.getItem(process.env.REACT_APP_TOKEN)}

const instance = axios.create({
  baseURL: baseURL,
  headers: {
    Authorization: `Bearer ${accessToken()}`,
    Accept: "application/json",
    "Content-Type": "application/json",
  },
});

instance.interceptors.request.use((config) => {
  const token = accessToken();
  if (token && config.headers) config.headers.Authorization = `Bearer ${token}`;
  config.validateStatus = (status) => status < 400;
  return config;
});

instance.interceptors.response.use(
  (successRes) => {
    return successRes;
  },
  (error) => {
    sweetalert({
      title: "Error!",
      text: error?.response?.data?.message
        ? error?.response?.data?.message
        : "An error occurred",
      icon: "error",
      buttons: {
        confirm: {
          text: "Close",
          value: true,
          visible: true,
          className: "btn bg-gradient-danger mx-auto",
          closeModal: true,
        },
      },
    });
    if (error.response.status == 401) {
      // localStorage.removeItem(process.env.REACT_APP_TOKEN);
      // window.location.reload()
    }
    console.log("caught error", error);
    return Promise.reject(error);
  }
);

const Api = {
  getUsers: async () => {
    return await instance.get(`/users`);
  },
  checkAdminRole: async () => {
    return await instance.get(`/me/memberOf`);
  },
  searchMessages: async (encodedSearchSubject, userId) => {
    console.log('encodedSearchSubject',encodedSearchSubject)
      return await instance.get(
        `/users/${userId}/messages?$search="subject:${encodedSearchSubject} OR from:${encodedSearchSubject}"`
      )
  },

 
  getFolders: async (userId = undefined) => {
    if (userId) {
      return await instance.get(`/users/${userId}/mailFolders`);
    }
    return await instance.get(`/me/mailFolders`);
  },
  moveToFolder: async ( messageId, payload, userId,) => {
      return await instance.post(
        `/users/${userId}/messages/${messageId}/move`,
        payload
      );
  },
};

export default Api;
