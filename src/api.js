import * as sweetalert from "sweetalert";
import axios from "axios";
import UserContext from "./context/UserContext";
import { useContext } from "react";

const baseURL = process.env.REACT_APP_API_URL;

const accessToken = localStorage.getItem(process.env.REACT_APP_TOKEN);

const instance = axios.create({
  baseURL: baseURL,
  headers: {
    Authorization: `Bearer ${accessToken}`,
    Accept: "application/json",
    "Content-Type": "application/json",
  },
});

instance.interceptors.request.use((config) => {
  const token = accessToken;
  if (token && config.headers) config.headers.Authorization = `Bearer ${token}`;
  config.validateStatus = (status) => status < 400;
  return config;
});

instance.interceptors.response.use(
  (successRes) => {
    console.log("successRes", successRes);
    // if (successRes?.data === 'Error: jwt expired') {
    //   localStorage.removeItem('lls-userinfo');
    // }
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
    if (error.response.data.error.code === "InvalidAuthenticationToken") {
      localStorage.removeItem(process.env.REACT_APP_TOKEN);
      const userContext = useContext(UserContext);
      userContext.update("");
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
  searchMessages: async (userId = undefined, encodedSearchSubject) => {
    if (userId) {
      return await instance.get(
        `/users/${userId}/messages?$search="subject:${encodedSearchSubject} OR from:${encodedSearchSubject}"`
      );
    }
    return await instance.get(
      `/me/messages?$search="subject:${encodedSearchSubject} OR from:${encodedSearchSubject}"`
    );
  },

  getMessages: async (userId = undefined, folderId) => {
    if (userId) {
      return await instance.get(
        `/users/${userId}/mailFolders/${folderId}/messages`
      );
    }
    return await instance.get(`/me/mailFolders/${folderId}/messages`);
  },
  getFolders: async (userId = undefined, folderId) => {
    if (userId) {
      return await instance.get(`/users/${userId}/mailFolders`);
    }
    return await instance.get(`/me/mailFolders`);
  },
  moveToFolder: async (userId = undefined, messageId, payload) => {
    if (userId) {
      return await instance.post(
        `/users/${userId}/messages/${messageId}/move`,
        payload
      );
    }
    return await instance.post(`/me/messages/${messageId}/move`, payload);
  },
};

export default Api;
