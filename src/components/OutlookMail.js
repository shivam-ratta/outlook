import React, { useEffect, useRef, useState } from "react";
import axios from "axios";
import moment from "moment";
import Api from "../api";
import * as sweetalert from "sweetalert";
import ReactPaginate from 'react-paginate';



const OutlookMail = () => {
  const [searchSubject, setSearchSubject] = useState("");
  const [messages, setMessages] = useState([]);
  const [isSearched, setIsSearched] = useState(false);
  const [token, setToken] = useState([]);
  const [timeFilter, setTimeFilter] = useState("");
  const [users, setUsers] = useState([]);
  const [selectAll, setSelectAll] = useState(false);
  const [loading, setLoading] = useState(0);
  const [page, setPage] = useState(0)
  const [limit, setLimit] = useState(50)


  const getUsers = async () => {
    Api.getUsers()
      .then((response) => {
        if (response?.users) {
          setUsers(response?.users);
        }
      })
      .catch((error) => {
        console.error("Error getting the list of users:", error);
      });
  };

  const handlePageClick = (event) => {
    setPage(event.selected)
  }

  const searchMessages = async () => {
    setPage(0)
    setLoading(true)
    setSelectedMessages([]);
    setSelectAll(false);
    try {
      setIsSearched(true);
      const encodedSearchSubject = encodeURIComponent(searchSubject);
      let promises = [];
      users.forEach((el) => {
        promises.push(
          Api.searchMessages(encodedSearchSubject, el?.id, el?.token)
            .then((res) => {
              // Use Promise.all to wait for both searchMessages and getFolders to complete
              return Promise.all([res, Api.getFolders(el?.id, el?.token)]);
            })
            .then(([res, folders]) => {
              // console.log("folders?.data?.", folders?.data);
              if (res?.data?.value) {
                // Use map to return an array of promises and then use Promise.all to wait for them
                const messages = res.data.value.map(async (message) => {
                  return {
                    token: el?.token,
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
            // console.log("filteredMessages", filteredMessages);
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
        }).finally(() => {
          setLoading(false)
        })

    } catch (error) {
      console.error("Error fetching messages:", error.message);
      setLoading(false)
    } finally {

    }
  };

  const moveToTrash = async () => {
    if (
      window.confirm("Are you sure you want to move these emails to trash?")
    ) {
      let movePromises;
      // if (isAdmin) {
      const messageIds = selectedMessages.map((el) => ({
        token: messages[el]?.token,
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
        return Api.moveToFolder(el.id, moveBody, el.userId, el?.token);
      });
      // } else {
      //   const messageIds = selectedMessages.map((el) => messages[el].id);
      //   const moveBody = {
      //     destinationId: trashFolderId,
      //   };

      //   movePromises = messageIds.map((messageId) => {
      //     return Api.moveToFolder(messageId, moveBody);
      //   });
      // }

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
      if (localStorage.getItem(process.env.REACT_APP_TOKEN) !== null) {
        setToken(JSON.parse(localStorage.getItem(process.env.REACT_APP_TOKEN)));
        getUsers();
      }
    }
  }, []);


  const logout = () => {
    localStorage.removeItem(process.env.REACT_APP_TOKEN);
    setToken("");
    setUsers([])
    setMessages([])
    setIsSearched(false)
    setTimeFilter('')
    setSearchSubject('')
    setPage(0)
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
    setLoading(true)
    const url = `${process.env.REACT_APP_BACKEND_API_URL}/login`;
    axios
      .get(url)
      .then((response) => {
        if (response?.data?.access_tokens) {
          localStorage.setItem(process.env.REACT_APP_TOKEN, JSON.stringify(response?.data?.access_tokens))
          setToken(response?.data?.access_tokens);
          setTimeout(() => {
            getUsers();
          }, 2000);
        }
      })
      .catch((error) => {
        console.error("Error:", error);
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
        })
      }).finally(() => {
        setLoading(false)
      });
  };

  return (
    <div>
      {token.length == 0 ? (
        <div
          className="d-flex align-items-center justify-content-center"
          style={{ height: "100vh" }}
        >
          <button disabled={loading} className="btn btn-primary" onClick={login}>
            {loading ? <>
              Logging in...  <i className="fa-solid fa-circle-notch fa-spin"></i>
            </> : 'Login'}
          </button>
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

                      <button
                        onClick={logout}
                        className="btn btn-danger me-2"
                      >
                        Logout
                      </button>
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
                {loading ? <div className="message-item border-bottom text-start p-3">
                  <div className="d-flex align-items-center justify-content-center">
                    <div className="col-12">
                      <div className="text-sm row">
                        <div className="col-12 text-center text-dark">
                          Loading...  <i className="fa-solid fa-circle-notch fa-spin"></i>
                        </div>
                      </div>
                    </div>
                  </div>
                </div> :
                  <>
                    {messages.length > 0 &&
                      <>
                      {messages.slice(page * limit, page * limit + limit).map((el, i) => (
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
                      { limit < messages?.length &&
                        <ReactPaginate
                            nextLabel={<i className="fa-solid fa-angle-right"></i>}
                            onPageChange={handlePageClick}
                            pageRangeDisplayed={5}
                            pageCount={Math.ceil(messages?.length / limit)}
                            previousLabel={<i className="fa-solid fa-angle-left"></i>}
                            renderOnZeroPageCount={null}
                            pageClassName="page-item"
                            pageLinkClassName="page-link"
                            previousClassName="page-item"
                            previousLinkClassName="page-link"
                            nextClassName="page-item"
                            nextLinkClassName="page-link"
                            breakLabel="..."
                            breakClassName="page-item"
                            breakLinkClassName="page-link"
                            containerClassName="pagination justify-content-center my-4"
                            activeClassName="active"
                          />
                      }
                      </>
                    }
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
                  </>
                }
              </div>
            </div>
          </div>
        </>
      )}
    </div>
  );
};

export default OutlookMail;
