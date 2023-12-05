import * as sweetalert from "sweetalert";
import axios from "axios";

const baseURL = process.env.REACT_APP_API_URL;

const accessTokens = () => {
  return JSON.parse(localStorage.getItem(process.env.REACT_APP_TOKEN));
};

const instance = axios.create({
  baseURL: baseURL,
  headers: {
    // Authorization: `Bearer `,
    Accept: "application/json",
    "Content-Type": "application/json",
  },
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
      localStorage.removeItem(process.env.REACT_APP_TOKEN);
      window.location.reload();
    }
    console.log("caught error", error);
    return Promise.reject(error);
  }
);

const Api = {
  getUsers: async () => {
    const tokens = accessTokens();
    let usersData = [];
    const userRequests = tokens.map(async (token) => {
      try {
        instance.defaults.headers.common["Authorization"] = `Bearer ${token}`;

        const response = await instance.get(`/users`);
        let users = response?.data?.value;
        // eslint-disable-next-line array-callback-return
        const usersTokens = users.map((user) => {
          user.token = token;
        });

        Promise.all(usersTokens)
          .then(() => {
            usersData = usersData.concat(response?.data?.value);
          })
          .catch((err) => {
            console.error(err);
          });
      } catch (error) {
        console.error(error);
      }
    });
    // return await instance.get(`/users`);
    return Promise.all(userRequests)
      .then(() => {
        return { users: usersData };
      })
      .catch((err) => {
        console.error(err);
      });
  },
  searchMessages: async (encodedSearchSubject, userId, token) => {
    instance.defaults.headers.common["Authorization"] = `Bearer ${token}`;
    return await instance.get(
      `/users/${userId}/messages?$search="subject:${encodedSearchSubject} OR from:${encodedSearchSubject}"`
    );
  },

  getFolders: async (userId, token) => {
    instance.defaults.headers.common["Authorization"] = `Bearer ${token}`;
    return await instance.get(`/users/${userId}/mailFolders`);
  },
  moveToFolder: async (messageId, payload, userId, token) => {
    instance.defaults.headers.common["Authorization"] = `Bearer ${token}`;
    return await instance.post(
      `/users/${userId}/messages/${messageId}/move`,
      payload
    );
  },
};

export default Api;
