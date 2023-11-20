import React, { useEffect, useState } from 'react';
import axios from 'axios';
import MicrosoftLogin from "react-microsoft-login";
import moment from 'moment';
import LoadingBar from 'react-top-loading-bar';


const OutlookMail = () => {
  // Replace these values with your app's credentials
  const clientId = process.env.REACT_APP_CLIENT_ID;
  const redirectUri = 'http://localhost:3000'; // Set this in your Azure AD App registration
  const [isLoggedIn, setIsLoggedIn] = useState(false)
  const [searchSubject, setSearchSubject] = useState('')
  const [messages, setMessages] = useState([])
  const [isSearched, setIsSearched] = useState(false)
  const [token, setToken] = useState('')
  const [timeFilter, setTimeFilter] = useState('')
  const [trashFolderId, setTranshFolderId] = useState('')
  const [selectedFolder, setSelectedFolder] = useState('')
  const [folders, setFolders] = useState([])
  const [users, setUsers] = useState([])
  const [selectedUser, setSelectedUser] = useState(null)
  const [selectAll, setSelectAll] = useState(false);
  const [isAdmin, setIsAdmin] = useState(false);
  const [loading, setLoading] = useState(0);

  const getFolder = async (accessToken) => {
    setSelectedFolder('')
    let folderUrl = isAdmin ? `https://graph.microsoft.com/v1.0/users/${selectedUser.id}/mailFolders` : 'https://graph.microsoft.com/v1.0/me/mailFolders';
    const headers = {
      'Authorization': `Bearer ${accessToken}`,
      'Accept': 'application/json',
    };
    axios.get(folderUrl, { headers }).then((res) => {
      if (res?.data?.value?.length) {
        setFolders(res.data.value)
          setSelectedFolder(res.data.value.find((el) => el?.displayName === "Inbox"))
        console.log('folders', res.data.value)
        let id = res.data.value.find((el) => el?.displayName === "Deleted Items").id
        setTranshFolderId(id)
        setLoading(100)
      }
    }).catch((err) => {
      console.log("err", err?.response?.data?.error)
    });
  }

  const getUsers = async (accessToken) => {
    const getUsersUrl = 'https://graph.microsoft.com/v1.0/users';
    const headers = {
      'Authorization': `Bearer ${accessToken}`,
      'Accept': 'application/json',
      'Content-Type': 'application/json',
    };
    axios.get(getUsersUrl, { headers })
      .then(response => {
        if (response?.data?.value) {
          console.log('List of users:', response.data.value);
          setUsers(response.data.value)
          let selectedUser = response.data.value.find((el) => el.id === localStorage.getItem(process.env.REACT_APP_USER_ID))
          // console.log('localStorage.getItem(process.env.REACT_APP_USER_ID)',localStorage.getItem(process.env.REACT_APP_USER_ID))
          // console.log('selectedUser',selectedUser)
          setSelectedUser(selectedUser)
        }
      })
      .catch(error => {
        console.error('Error getting the list of users:', error.message);
        console.log('Response Status:', error.response.status);
        console.log('Response Data:', error.response.data);
      });
  }

  // Get user's messages from Outlook
  const getMessages = async (accessToken) => {
    setSelectedMessages([])
    setSelectAll(false)
    if (!accessToken) {
      return
    }
    try {
      setIsSearched(false)
      let messagesUrl = isAdmin ? `https://graph.microsoft.com/v1.0/users/${selectedUser.id}/mailFolders/${selectedFolder.id}/messages` : `https://graph.microsoft.com/v1.0/me/mailFolders/${selectedFolder.id}/messages`;
      const headers = {
        'Authorization': `Bearer ${accessToken}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json',
      };

      await axios.get(messagesUrl, { headers }).then((res) => {
        if (res?.data?.value) {
          if (timeFilter.length) {
            let filter = ''
            if (timeFilter === '24hours') {
              filter = moment().subtract(24, 'hours').valueOf()
            } else if (timeFilter === '48hours') {
              filter = moment().subtract(48, 'hours').valueOf()
            } else if (timeFilter === '1week') {
              filter = moment().subtract(1, 'weeks').valueOf()
            } else if (timeFilter === '2week') {
              filter = moment().subtract(2, 'weeks').valueOf()
            } else if (timeFilter === '1month') {
              filter = moment().subtract(1, 'months').valueOf()
            };
            const filterData = res.data.value.filter((el) => moment(el.receivedDateTime).valueOf() >= filter && moment(el.receivedDateTime).valueOf() <= moment().valueOf())
            setMessages(filterData)
            return
          }
          setMessages(res.data.value)
        }
      }).catch((err) => {
        console.log("err", err.response.data.error)
        if (err.response.data.error.code === 'InvalidAuthenticationToken') {
          logout()
        }
      });

    } catch (error) {
      console.error('Error fetching messages:', error.message);
    }
  };

  const searchMessages = async (accessToken) => {
    setSelectedMessages([])
    setSelectAll(false)
    if (!accessToken) {
      return
    }
    try {
      setIsSearched(true)
      let messagesUrl = isAdmin ? `https://graph.microsoft.com/v1.0/users/${selectedUser.id}/messages` : `https://graph.microsoft.com/v1.0/me/messages`;

      if (searchSubject.length) {
        const encodedSearchSubject = encodeURIComponent(searchSubject);
        messagesUrl = `${messagesUrl}?$search="subject:${encodedSearchSubject} OR from:${encodedSearchSubject}"`;
      }

      // if (searchSubject.length) {
      //   const encodedSearchSubject = encodeURIComponent(searchSubject);
      //   messagesUrl = `https://graph.microsoft.com/v1.0/users/messages?$search="mail:${encodedSearchSubject}"`;
      // }


      const headers = {
        'Authorization': `Bearer ${accessToken}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'ConsistencyLevel': 'eventual',
      };

      await axios.get(messagesUrl, { headers }).then((res) => {
        if (res?.data?.value) {
          if (timeFilter.length) {
            let filter = ''
            if (timeFilter === '24hours') {
              filter = moment().subtract(24, 'hours').valueOf()
            } else if (timeFilter === '48hours') {
              filter = moment().subtract(48, 'hours').valueOf()
            } else if (timeFilter === '1week') {
              filter = moment().subtract(1, 'weeks').valueOf()
            } else if (timeFilter === '2week') {
              filter = moment().subtract(2, 'weeks').valueOf()
            } else if (timeFilter === '1month') {
              filter = moment().subtract(1, 'months').valueOf()
            };
            const filterData = res.data.value.filter((el) => moment(el.receivedDateTime).valueOf() >= filter && moment(el.receivedDateTime).valueOf() <= moment().valueOf())
            setMessages(filterData)
            return
          }
          setMessages(res.data.value)
        }
      }).catch((err) => {
        console.log("err", err.response.data.error)
        if (err.response.data.error.code === 'InvalidAuthenticationToken') {
          logout()
        }
      });

    } catch (error) {
      console.error('Error fetching messages:', error.message);
    }
  };

  const moveToTrash = (accessToken) => {
    if (window.confirm('Are you sure you want to these emails to trash?')) {
      const messageIds = selectedMessages.map((el) => messages[el].id);

      const moveBody = {
        destinationId: trashFolderId,
      };

      // Set up Axios headers
      const moveHeaders = {
        'Authorization': `Bearer ${accessToken}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json',
      };

      const movePromises = messageIds.map((messageId) => {
        const moveMessagesUrl = isAdmin ? `https://graph.microsoft.com/v1.0/users/${selectedUser.id}/messages/${messageId}/move` : `https://graph.microsoft.com/v1.0/me/messages/${messageId}/move`;
        return axios.post(moveMessagesUrl, moveBody, { headers: moveHeaders });
      });

      axios.all(movePromises)
        .then(axios.spread((...responses) => {
          console.log('Messages moved to Trash:', responses);

          responses.forEach((response, index) => {
            console.log(`Response for message ${messageIds[index]}:`, response.data);
          });

          // Perform additional actions, e.g., fetching updated messages
          getMessages(token);
          setSelectAll(false);
          setSelectedMessages([]);
        }))
        .catch(errors => {
          // Handle errors here
          console.error('Error moving messages to Trash:', errors);

          errors?.length && errors?.forEach((error, index) => {
            console.log(`Error for message ${messageIds[index]}:`, error.message);
            console.log(`Response Status for message ${messageIds[index]}:`, error.response.status);
            console.log(`Response Data for message ${messageIds[index]}:`, error.response.data);
          });
        });
    }
  };

  const moveToFolder = (id, accessToken) => {
    if (window.confirm('Are you sure you want to these emails?')) {
      const messageIds = selectedMessages.map((el) => messages[el].id);

      const moveBody = {
        destinationId: id,
      };

      // Set up Axios headers
      const moveHeaders = {
        'Authorization': `Bearer ${accessToken}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json',
      };

      const movePromises = messageIds.map((messageId) => {
        const moveMessagesUrl = isAdmin ? `https://graph.microsoft.com/v1.0/users/${selectedUser.id}/messages/${messageId}/move` : `https://graph.microsoft.com/v1.0/me/messages/${messageId}/move`;
        return axios.post(moveMessagesUrl, moveBody, { headers: moveHeaders });
      });

      axios.all(movePromises)
        .then(axios.spread((...responses) => {
          console.log('Messages moved to Trash:', responses);

          responses.forEach((response, index) => {
            console.log(`Response for message ${messageIds[index]}:`, response.data);
          });

          // Perform additional actions, e.g., fetching updated messages
          getMessages(token);
          setSelectAll(false);
          setSelectedMessages([]);
        }))
        .catch(errors => {
          // Handle errors here
          console.error('Error moving messages to Trash:', errors);

          errors?.length  && errors?.forEach((error, index) => {
            console.log(`Error for message ${messageIds[index]}:`, error.message);
            console.log(`Response Status for message ${messageIds[index]}:`, error.response.status);
            console.log(`Response Data for message ${messageIds[index]}:`, error.response.data);
          });
        });
    }
  }


  useEffect(() => {
    if (!!localStorage.getItem(process.env.REACT_APP_TOKEN)) {
      getUsers(localStorage.getItem(process.env.REACT_APP_TOKEN))
      setToken(localStorage.getItem(process.env.REACT_APP_TOKEN));
      checkAdminRole(localStorage.getItem(process.env.REACT_APP_TOKEN))
    } else {
      setToken('')
    }
  }, [])

  const authHandler = (err, data) => {
    console.log('user data', data)
    if (data?.accessToken) {
      localStorage.setItem(process.env.REACT_APP_TOKEN, data.accessToken)
      localStorage.setItem(process.env.REACT_APP_USER_ID, data.uniqueId)
      setToken(data.accessToken);
      setIsLoggedIn(true)
      checkAdminRole(data.accessToken);
      getUsers(data.accessToken)
    }
  };

  const checkAdminRole = (accessToken) => {
    const getUsersUrl = 'https://graph.microsoft.com/v1.0/me/memberOf';
    const headers = {
      'Authorization': `Bearer ${accessToken}`,
      'Accept': 'application/json',
      'Content-Type': 'application/json',
    };
    axios.get(getUsersUrl, { headers })
      .then(response => {
        console.log('checkAdminRole', response?.data);
        if (response?.data.value?.length && response?.data.value[0].displayName === "Global Administrator") {
          setIsAdmin(true)
        } else {
          setIsAdmin(false)
        }
      })
      .catch(error => {
        console.error('Error getting the list of users:', error.message);
        console.log('Response Status:', error.response.status);
        console.log('Response Data:', error.response.data);
      });
  };

  const logout = () => {
    localStorage.removeItem(process.env.REACT_APP_TOKEN)
    setToken('')
    setIsLoggedIn(false)
  }

  useEffect(() => {
    if (selectedUser?.id) {
      setSelectedMessages([])
      setSelectAll(false)
      getFolder(token)
    }
  }, [selectedUser])

  useEffect(() => {
    if(selectedFolder) {
      console.log('selectedFolder',selectedFolder)
      setSelectedMessages([])
      setSelectAll(false)
      setIsSearched(false)
      getMessages(token)
      setSearchSubject('')
      setTimeFilter('')

      // if (isSearched) {
      //   searchMessages(token)
      // } else {
      //   getMessages(token)
      // }
    }
  }, [selectedFolder])

  useEffect(() => {
    if (searchSubject === '') {
      getMessages(token)
    }
  }, [searchSubject])

  useEffect(() => {
    if (isSearched) {
      searchMessages(token)
    } else {
      getMessages(token)
    }
  }, [timeFilter])



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

  const handleUserChange = (e) => {
    const selectedUserId = e.target.value;
    setSelectedUser(users.find((el) => el.id === selectedUserId));
  };


  return (
    <div>
      {
        token === '' ?
          <div className='d-flex align-items-center justify-content-center' style={{ height: '100vh' }}>
            <MicrosoftLogin clientId={clientId} authCallback={authHandler} />
          </div> :
          <>
           <LoadingBar
        color="#0d6efd"
        progress={loading}
        onLoaderFinished={() => setLoading(0)}
      />
            <div className="container-fluid px-0">
              <div className="d-flex">
                <div className="col-lg-2 d-flex align-items-center justify-content-center"><div className="fw-bold text-primary me-2 my-lg-2 my-3 fs-4"> Outlook</div></div>
                <div className="col-lg-10 px-0">
                  <div className="d-flex p-3 align-items-center justify-content-center">
                    {/* Left Section */}
                    <div className="col-1 ">
                     <div className='d-flex  align-items-center justify-content-center'>
                        <div className="col-1 d-flex align-items-center justify-content-center">
                          {messages.length > 0 && <div className="form-check">
                            <input className="form-check-input" checked={selectedMessages.length === messages.length && selectedMessages.length > 0} onChange={(e) => e.target.checked ? handleSelectAll() : handleUnselectAll()} type="checkbox" value="" id={`flexCheckDefaultAll`} />
                          </div>}
                        </div>
                      </div>
                    </div>
                    {/* Right Section */}
                    <div className="col-11  d-lg-flex align-items-center justify-content-between">

                      <div className="fw-bold text-primary me-2 my-lg-2 my-3">{isSearched ? 'Search' : selectedFolder?.displayName ? selectedFolder?.displayName : ''}</div>
                      <div className="d-lg-flex align-items-center justify-content-end">
                        <div className="me-2 my-2 d-inline-block position-relative">
                          <input
                            type="email"
                            className="form-control searchInput mx-auto"
                            id="exampleFormControlInput1"
                            placeholder="Search by email and Subject"
                            value={searchSubject}
                            onInput={(e) => { setSearchSubject(e.target.value) }}
                          />

                          {isSearched && (
                            <div className="col-lg-2">
                              <button onClick={() => { setSearchSubject(''); }} className='btn position-absolute close-search  me-2'>
                                <i className='fa fa-times'></i>
                              </button>
                            </div>
                          )}
                        </div>

                        <div className="d-inline-block my-2 ">
                          <button onClick={() => searchMessages(token)} className='btn btn-primary me-2'>
                            Search
                          </button>
                        </div>
                        <div className="d-inline-block my-2 ">
                          <button onClick={() => isSearched ? searchMessages(token) : getMessages(token)} className='btn btn-primary me-2'>
                            <i className="fa fa-refresh me-2" aria-hidden="true"></i>
                            Refresh
                          </button>
                        </div>
                        {selectedMessages.length > 0 && selectedFolder.displayName !== 'Deleted Items' && <div className="d-inline-block my-2 ">
                          <button onClick={() => moveToTrash(token)} className='btn btn-danger me-2'>
                            <i className='fa fa-trash me-2'></i> Trash
                          </button>
                        </div>}
                        {selectedMessages.length > 0 && selectedFolder.displayName == 'Deleted Items' && <div className="d-inline-block me-2 floatingSelect my-2 ">
                        <select className="form-select mb-0" onChange={(e) => { moveToFolder(e.target.value, token) }} id="floatingSelect" aria-label="Floating label select example">
                            <option value="">Move to</option>
                            {folders.length > 0 && folders.map((el, i) => (
                            <option value={el?.id} key={i}>{el?.displayName}</option>
                            ))}
                          </select>
                        </div>}
                        <div className="d-inline-block my-2 ">
                          <button onClick={logout} className='btn btn-danger me-2'>
                            Logout
                          </button>
                        </div>

                        <div className="d-inline-block my-2 me-2">
                          <select className="form-select floatingSelect mb-0" onChange={(e) => { setTimeFilter(e.target.value) }} id="" aria-label="Floating label select example">
                            <option value="">Filter</option>
                            <option value="24hours">24 Hours</option>
                            <option value="48hours">48 Hours</option>
                            <option value="1week">1 Week</option>
                            <option value="2week">2 Weeks</option>
                            <option value="1month">1 Month</option>
                          </select>
                        </div>

                        {isAdmin &&
                          <div className="dropdown">
                            <button className="btn border-secondary rounded-circle user-icon-btn dropdown-toggle" type="button" id="dropdownMenuButton1" data-bs-toggle="dropdown" aria-expanded="false">
                              <i className='fa fa-user'></i>
                            </button>
                            <ul className="dropdown-menu" aria-labelledby="dropdownMenuButton1">
                              <li className="dropdown-item">Name:{selectedUser?.displayName}</li>
                              <li className="dropdown-item"> Email: {selectedUser?.mail}</li>
                              <li className="dropdown-item d-flex align-items-center w-100">
                                <div className='me-2'>Change User:</div>
                                <div className=''>
                                  <select key={selectedUser} defaultValue={selectedUser?.id} className="form-select w-100 mb-0" onChange={(e) => { handleUserChange(e) }} id="floatingSelect" aria-label="Floating label select example">
                                    <option value="" disabled>Select User</option>
                                    {users.map((el, i) => {
                                      return (<option key={i} value={el.id}>{el?.mail}</option>)
                                    })}
                                  </select>
                                </div>
                              </li>
                            </ul>
                          </div>}
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              {/* Messages Section */}
              <div className="w-100 d-flex px-0 bg-light shadow" style={{ minHeight: '100vh' }}>
                <div className='col-2 px-2 border-top border-end'>
                  {folders.length > 0 && folders.map((el, i) => (
                    <div onClick={() => { setSelectedFolder(el) }} className={`message-item cursor-pointer text-start p-3 mb-0 align-items-center justify-content-center ${selectedFolder?.displayName === el?.displayName && !isSearched && 'alert alert-primary'}`} key={i}>
                      <div className="text-sm cursor-pointer d-flex align-items-center">
                        <i className='far fa-folder me-3'></i>
                        <div>{el?.displayName} </div>
                      </div>
                    </div>
                  ))} </div>
                <div className='col-10'>
                  {messages.length > 0 && messages.map((el, i) => (
                    <div className="message-item border-bottom text-start p-3 d-flex align-items-center justify-content-center" key={i}>
                      <div className="col-1 d-flex align-items-center justify-content-center">
                        <div className="form-check">
                          <input className="form-check-input" checked={selectAll || selectedMessages.includes(i)} onChange={() => handleCheckboxChange(i)} type="checkbox" value="" id={`flexCheckDefault${i}`} />
                        </div>
                      </div>
                      <div className="col-11">
                        <div className="text-sm row">
                          <div className="col-6">{el?.sender?.emailAddress?.name || 'No name'} </div>
                          <div className='text-primary col-6 text-end text-sm'>{moment(el?.receivedDateTime).format('lll')}</div>
                        </div>
                        <div className="text-sm text-primary">{el?.subject || 'No Subject'}</div>
                        <div className="text-sm text-ellipsis">{el?.bodyPreview || 'No body'}</div>
                      </div>
                    </div>
                  ))}
                  {messages.length === 0 && <div className="message-item border-bottom text-start p-3">
                    <div className="d-flex align-items-center justify-content-center">
                      <div className="col-12">
                        <div className="text-sm row">
                          <div className="col-12 text-center">No Emails Found</div>
                        </div>
                      </div>
                    </div>
                  </div>
                  }
                  {/* {messages.length == 0 && !isSearched && <div className="message-item border-bottom text-start p-3">
                    <div className="d-flex align-items-center justify-content-center">
                      <div className="col-12">
                        <div className="text-sm row">
                          <div className="col-12 text-center">Search Email</div>
                        </div>
                      </div>
                    </div>
                  </div>
                  } */}
                </div>
              </div>
            </div>


          </>
      }

    </div>
  );
};

export default OutlookMail;
