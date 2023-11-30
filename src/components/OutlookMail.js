import React, { useEffect, useRef, useState } from "react";
import axios from "axios";
import MicrosoftLogin from "react-microsoft-login";
import moment from "moment";
import Api from "../api";
import * as sweetalert from "sweetalert";

const OutlookMail = () => {
  // Replace these values with your app's credentials
  const clientId = process.env.REACT_APP_CLIENT_ID;
  const [searchSubject, setSearchSubject] = useState("");
  const [messages, setMessages] = useState([]);
  const [isSearched, setIsSearched] = useState(false);
  const [token, setToken] = useState("");
  const [timeFilter, setTimeFilter] = useState("");
  const [trashFolderId, setTranshFolderId] = useState("");
  const [selectedFolder, setSelectedFolder] = useState("");
  const [folders, setFolders] = useState([]);
  const [users, setUsers] = useState([]);
  const [selectedUser, setSelectedUser] = useState(null);
  const [selectAll, setSelectAll] = useState(false);
  const [isAdmin, setIsAdmin] = useState(false);
  const [loading, setLoading] = useState(0);

  const getTrashFolder = () => {
    Api.getFolders().then((res) => {
      if (res?.data?.value) {
        setTranshFolderId(
          res.data.value.find((el) => el.displayName == "Deleted Items")?.id
        );
      }
    });
  };

  const getUsers = async () => {
    Api.getUsers()
      .then((response) => {
        if (response?.data?.value) {
          console.log("List of users:", response.data.value);
          setUsers(response.data.value);
          let selectedUser = response.data.value.find(
            (el) =>
              el.id === localStorage.getItem(process.env.REACT_APP_USER_ID)
          );
          // console.log('localStorage.getItem(process.env.REACT_APP_USER_ID)',localStorage.getItem(process.env.REACT_APP_USER_ID))
          // console.log('selectedUser',selectedUser)
          setSelectedUser(selectedUser);
        }
      })
      .catch((error) => {
        console.error("Error getting the list of users:", error);
      });
  };

  const searchMessages = async () => {
    setSelectedMessages([]);
    setSelectAll(false);
    try {
      setIsSearched(true);
      const encodedSearchSubject = encodeURIComponent(searchSubject);
      let promise;
        let promises = [];
        users.forEach((el) => {
          promises.push(
            Api.searchMessages(encodedSearchSubject, el?.id)
              .then((res) => {
                // Use Promise.all to wait for both searchMessages and getFolders to complete
                return Promise.all([res, Api.getFolders(el?.id)]);
              })
              .then(([res, folders]) => {
                console.log("folders?.data?.", folders?.data);
                if (res?.data?.value) {
                  // Use map to return an array of promises and then use Promise.all to wait for them
                  const messages = res.data.value.map(async (message) => {
                    return {
                      userId: el.id,
                      ...message,
                      folders: folders?.data?.value ? folders?.data?.value : [],
                    };
                  });
                  return Promise.all(messages);
                }
                return [];
              })
              .catch((error) => {
                console.error(error);
                return []; // Handle errors by returning an empty array
              })
          );
        });

        Promise.all(promises)
          .then((userMessages) => {
            // Flatten the array of arrays into a single array of messages
            const flatMessages = userMessages.flat();

            if (flatMessages.length) {
              let filteredMessages = flatMessages;
              console.log("filteredMessages", filteredMessages);
              if (timeFilter.length) {
                let filter = "";

                // Use a switch statement for better readability
                switch (timeFilter) {
                  case "24hours":
                    filter = moment().subtract(24, "hours").valueOf();
                    break;
                  case "48hours":
                    filter = moment().subtract(48, "hours").valueOf();
                    break;
                  case "1week":
                    filter = moment().subtract(1, "weeks").valueOf();
                    break;
                  case "2week":
                    filter = moment().subtract(2, "weeks").valueOf();
                    break;
                  case "1month":
                    filter = moment().subtract(1, "months").valueOf();
                    break;
                  default:
                    break;
                }

                filteredMessages = filteredMessages.filter(
                  (el) =>
                    moment(el.receivedDateTime).valueOf() >= filter &&
                    moment(el.receivedDateTime).valueOf() <= moment().valueOf()
                );
              }

              setMessages(filteredMessages);
            } else {
              setMessages([]);
            }
          })
          .catch((error) => {
            console.error(error);
          });
      
    } catch (error) {
      console.error("Error fetching messages:", error.message);
    }
  };

  const moveToTrash = async () => {
    if (
      window.confirm("Are you sure you want to move these emails to trash?")
    ) {
      let movePromises;
      if (isAdmin) {
        const messageIds = selectedMessages.map((el) => ({
          userId: messages[el].userId,
          id: messages[el].id,
          folderId: messages[el]?.folders.find(
            (folder) => folder?.displayName === "Deleted Items"
          )?.id,
        }));

        movePromises = messageIds.map((el) => {
          const moveBody = {
            destinationId: el.folderId,
          };
          return Api.moveToFolder(el.id, moveBody, el.userId);
        });
      } else {
        const messageIds = selectedMessages.map((el) => messages[el].id);
        const moveBody = {
          destinationId: trashFolderId,
        };

        movePromises = messageIds.map((messageId) => {
          return Api.moveToFolder(messageId, moveBody);
        });
      }

      await axios
        .all(movePromises)
        .then(() => {
          setSelectAll(false);
          setSelectedMessages([]);
          searchMessages();
          sweetalert({
            title: "Success",
            text: "Emails Moved Successfully",
            icon: "success",
            buttons: {
              confirm: {
                text: "Ok",
                value: true,
                visible: true,
                className: "btn bg-gradient-success mx-auto",
                closeModal: true,
              },
            },
          });
        })
        .catch((errors) => {
          // Handle errors here
          console.error("Error moving messages to Trash:", errors);
        });
    }
  };

  let mounted = useRef(null);
  useEffect(() => {
    if (!mounted.current) {
      mounted.current = true;
      if (localStorage.getItem(process.env.REACT_APP_TOKEN)) {
        setToken(localStorage.getItem(process.env.REACT_APP_TOKEN));
        getUsers();
        // checkAdminRole();
      }
    }
  }, []);

  const authHandler = (err, data) => {
    console.log("user data", data);
    if (data?.accessToken) {
      localStorage.setItem(process.env.REACT_APP_TOKEN, data.accessToken);
      localStorage.setItem(process.env.REACT_APP_USER_ID, data.uniqueId);
      setToken(data?.accessToken);
      setTimeout(() => {
        // checkAdminRole();
        getUsers();
      }, 2000);
    }
  };

  const checkAdminRole = () => {
    Api.checkAdminRole()
      .then((response) => {
        if (
          response?.data.value?.length &&
          response?.data.value[0].displayName === "Global Administrator"
        ) {
          setIsAdmin(true);
        } else {
          setIsAdmin(false);
          getTrashFolder();
        }
      })
      .catch((error) => {
        console.error("Error getting the list of users:", error);
      });
  };

  const logout = () => {
    localStorage.removeItem(process.env.REACT_APP_TOKEN);
    setToken("");
  };

  useEffect(() => {
    if (isSearched && token) {
      searchMessages();
    }
  }, [timeFilter]);

  const [selectedMessages, setSelectedMessages] = useState([]);

  const handleCheckboxChange = (index) => {
    const isSelected = selectedMessages.includes(index);
    if (isSelected) {
      setSelectedMessages((prevSelected) =>
        prevSelected.filter((item) => item !== index)
      );
    } else {
      setSelectedMessages((prevSelected) => [...prevSelected, index]);
    }
  };

  const handleSelectAll = () => {
    // Select all checkboxes
    setSelectedMessages(Array.from({ length: messages.length }, (_, i) => i));
    setSelectAll(true);
  };

  const handleUnselectAll = () => {
    // Unselect all checkboxes
    setSelectedMessages([]);
    setSelectAll(false);
  };

  const login = () => {
    const clientId = process.env.REACT_APP_CLIENT_ID;
    const tenantId = process.env.REACT_APP_TENANT_ID;
    const clientSecret = process.env.REACT_APP_CLIENT_SECRET;

    const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

    const data = new URLSearchParams();
    data.append("client_id", clientId);
    data.append("scope", "https://graph.microsoft.com/.default");
    data.append("client_secret", clientSecret);
    data.append("grant_type", "client_credentials");

    axios
      .post(url, data, {
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
      })
      .then((response) => {
        console.log("Access Token:", response.data.access_token);
      })
      .catch((error) => {
        console.error("Error:", error);
      });

      
  };

  return (
    <div>
      {token === "" ? (
        <div
          className="d-flex align-items-center justify-content-center"
          style={{ height: "100vh" }}
        >
          {/* <MicrosoftLogin clientId={clientId} authCallback={authHandler} /> */}
          <button className="btn btn-primary" onClick={login}>login</button>
        </div>
      ) : (
        <>
          <div className="container-fluid px-0">
            <div className="d-flex">
              <div className="col-lg-12 px-0">
                <div className="d-flex p-3 align-items-center justify-content-center">
                  {/* Left Section */}
                  <div className="col-1 ">
                    <div className="d-flex  align-items-center justify-content-center">
                      <div className="col-1 d-flex align-items-center justify-content-center">
                        {messages.length > 0 && (
                          <div className="form-check">
                            <input
                              className="form-check-input"
                              checked={
                                selectedMessages.length === messages.length &&
                                selectedMessages.length > 0
                              }
                              onChange={(e) =>
                                e.target.checked
                                  ? handleSelectAll()
                                  : handleUnselectAll()
                              }
                              type="checkbox"
                              value=""
                              id={`flexCheckDefaultAll`}
                            />
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                  {/* Right Section */}
                  <div className="col-11  d-lg-flex align-items-center justify-content-between">
                    <div className="fw-bold text-primary me-2 my-lg-2 my-3">
                      Outlook Mails
                    </div>
                    <div className="d-lg-flex align-items-center justify-content-end">
                      <div className="me-2 my-2 d-inline-block position-relative">
                        <input
                          type="email"
                          className="form-control searchInput mx-auto"
                          id="exampleFormControlInput1"
                          placeholder="Search by email and Subject"
                          value={searchSubject}
                          onInput={(e) => {
                            setSearchSubject(e.target.value);
                          }}
                        />

                        {isSearched && (
                          <div className="col-lg-2">
                            <button
                              onClick={() => {
                                setSearchSubject("");
                              }}
                              className="btn position-absolute close-search  me-2"
                            >
                              <i className="fa fa-times"></i>
                            </button>
                          </div>
                        )}
                      </div>

                      <div className="d-inline-block my-2 ">
                        <button
                          onClick={() => searchMessages()}
                          className="btn btn-primary me-2"
                        >
                          Search
                        </button>
                      </div>
                      {selectedMessages.length > 0 && (
                        <div className="d-inline-block my-2 ">
                          <button
                            onClick={() => moveToTrash()}
                            className="btn btn-danger me-2"
                          >
                            <i className="fa fa-trash me-2"></i> Trash
                          </button>
                        </div>
                      )}

                      <div className="d-inline-block my-2 me-2">
                        <select
                          className="form-select floatingSelect mb-0"
                          onChange={(e) => {
                            setTimeFilter(e.target.value);
                          }}
                          id=""
                          aria-label="Floating label select example"
                        >
                          <option value="">Filter</option>
                          <option value="24hours">24 Hours</option>
                          <option value="48hours">48 Hours</option>
                          <option value="1week">1 Week</option>
                          <option value="2week">2 Weeks</option>
                          <option value="1month">1 Month</option>
                        </select>
                      </div>

                      <div className="dropdown">
                        <button
                          className="btn border-secondary rounded-circle user-icon-btn dropdown-toggle"
                          type="button"
                          id="dropdownMenuButton1"
                          data-bs-toggle="dropdown"
                          aria-expanded="false"
                        >
                          <i className="fa fa-user"></i>
                        </button>
                        <ul
                          className="dropdown-menu"
                          aria-labelledby="dropdownMenuButton1"
                        >
                          <li className="dropdown-item">
                            Name:{selectedUser?.displayName}
                          </li>
                          <li className="dropdown-item">
                            {" "}
                            Email: {selectedUser?.mail}
                          </li>
                          <li className="dropdown-item">
                            <button
                              onClick={logout}
                              className="btn btn-danger me-2"
                            >
                              Logout
                            </button>
                          </li>
                        </ul>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            {/* Messages Section */}
            <div
              className="w-100 d-flex px-0 bg-light shadow"
              style={{ minHeight: "100vh" }}
            >
              <div className="col-12">
                {messages.length > 0 &&
                  messages.map((el, i) => (
                    <div
                      className="message-item border-bottom text-start p-3 d-flex align-items-center justify-content-center"
                      key={i}
                    >
                      <div className="col-1 d-flex align-items-center justify-content-center">
                        <div className="form-check">
                          <input
                            className="form-check-input"
                            checked={selectAll || selectedMessages.includes(i)}
                            onChange={() => handleCheckboxChange(i)}
                            type="checkbox"
                            value=""
                            id={`flexCheckDefault${i}`}
                          />
                        </div>
                      </div>
                      <div className="col-11">
                        <div className="text-sm row">
                          <div className="col-6">
                            {el?.sender?.emailAddress?.name || "No name"}
                          </div>
                          <div className="text-primary col-6 text-end text-sm">
                            {moment(el?.receivedDateTime).format("lll")}
                          </div>
                        </div>
                        <div className="text-sm text-primary">
                          {el?.subject || "No Subject"}
                        </div>
                        <div className="text-sm text-ellipsis">
                          {el?.bodyPreview || "No body"}
                        </div>
                      </div>
                    </div>
                  ))}
                {messages.length === 0 && (
                  <div className="message-item border-bottom text-start p-3">
                    <div className="d-flex align-items-center justify-content-center">
                      <div className="col-12">
                        <div className="text-sm row">
                          <div className="col-12 text-center">
                            No Emails Found
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        </>
      )}
    </div>
  );
};

export default OutlookMail;
