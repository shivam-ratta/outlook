import React, { useEffect, useState } from 'react';
import axios from 'axios';
import MicrosoftLogin from "react-microsoft-login";
import moment from 'moment';

const OutlookMail = () => {
  // Replace these values with your app's credentials
  const clientId = process.env.REACT_APP_CLIENT_ID;
  const redirectUri = 'http://localhost:3000'; // Set this in your Azure AD App registration
  const tenantId = process.env.REACT_APP_TENANT_ID;
  const [isLoggedIn, setIsLoggedIn] = useState(false)
  const [searchSubject, setSearchSubject] = useState('')
  const [messages, setMessages] = useState([])
  const [isSearched, setIsSearched] = useState('')
  const [token, setToken] = useState('')
  const [timeFilter, setTimeFilter] = useState('')
  const [trashFolderId, setTranshFolderId] = useState('')
  const [selectAll, setSelectAll] = useState(false);


  // Azure AD endpoint for authorization
  const authorizeUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`;

  // Get the authorization code
  const getAuthorizationCode = () => {
    const queryParams = new URLSearchParams({
      client_id: clientId,
      response_type: 'code',
      redirect_uri: redirectUri,
      response_mode: 'query',
      scope: 'offline_access openid user.read mail.read',
    });

    window.location.href = `${authorizeUrl}?${queryParams.toString()}`;
  };

  // Exchange the authorization code for an access token
  const getAccessToken = async (code) => {
    const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    const tokenData = {
      grant_type: 'authorization_code',
      client_id: clientId,
      code,
      redirect_uri: redirectUri,
      scope: 'offline_access openid user.read mail.read',
    };

    try {
      const tokenResponse = await axios.post(tokenUrl, new URLSearchParams(tokenData), {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
        },
      });

      return tokenResponse.data.access_token;
    } catch (error) {
      console.error('Error getting access token:', error.message);
      throw error;
    }
  };

  const getFolder = async (accessToken) => {
    let folderUrl = 'https://graph.microsoft.com/v1.0/me/mailFolders';
    const headers = {
      'Authorization': `Bearer ${accessToken}`,
      'Accept': 'application/json',
    };
    axios.get(folderUrl, { headers }).then((res) => {
      if (res?.data?.value) {
        let id = res.data.value.find((el) => el.displayName == "Deleted Items").id
        // console.log('id', id)
        setTranshFolderId(id)
      }
    }).catch((err) => {
      console.log("err", err.response.data.error)
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
      getFolder(accessToken)
      let messagesUrl = 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages';
      if (searchSubject.length) {
        const encodedSearchSubject = encodeURIComponent(searchSubject);
        const filterDate = moment().subtract(timeFilter, 'hours').toISOString();
        ;

        messagesUrl = `${messagesUrl}?$search="subject:${encodedSearchSubject} OR from:${encodedSearchSubject}"`;
        setIsSearched(true);
      }


      const headers = {
        'Authorization': `Bearer ${accessToken}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json',
      };

      await axios.get(messagesUrl, { headers }).then((res) => {
        if (res?.data?.value) {
          console.log('res.data.value', res.data.value)
          if (timeFilter.length) {
            let filter = ''
            if (timeFilter == '24hours') {
              filter = moment().subtract(24, 'hours').valueOf()
            } else if (timeFilter == '48hours') {
              filter = moment().subtract(48, 'hours').valueOf()
            } else if (timeFilter == '1week') {
              filter = moment().subtract(1, 'weeks').valueOf()
            } else if (timeFilter == '2week') {
              filter = moment().subtract(2, 'weeks').valueOf()
            } else if (timeFilter == '1month') {
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
        if (err.response.data.error.code == 'InvalidAuthenticationToken') {
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
        const moveMessagesUrl = `https://graph.microsoft.com/v1.0/me/messages/${messageId}/move`;
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

          errors.forEach((error, index) => {
            console.log(`Error for message ${messageIds[index]}:`, error.message);
            console.log(`Response Status for message ${messageIds[index]}:`, error.response.status);
            console.log(`Response Data for message ${messageIds[index]}:`, error.response.data);
          });
        });
    }
  };


  useEffect(() => {
    if (!!localStorage.getItem(process.env.REACT_APP_TOKEN)) {
      getMessages(localStorage.getItem(process.env.REACT_APP_TOKEN));
      setToken(localStorage.getItem(process.env.REACT_APP_TOKEN));
    } else {
      setToken('')
    }
  }, [])

  const authHandler = (err, data) => {
    if (data?.accessToken) {
      localStorage.setItem(process.env.REACT_APP_TOKEN, data.accessToken)
      setToken(data.accessToken);
      setIsLoggedIn(true)
      getMessages(data.accessToken)
    }
  };

  const logout = () => {
    localStorage.removeItem(process.env.REACT_APP_TOKEN)
    setToken('')
    setIsLoggedIn(false)
  }

  useEffect(() => {
    if (searchSubject == '' && isSearched) {
      getMessages(token)
    }
  }, [searchSubject])

  useEffect(() => {
    getMessages(token)
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

  return (
    <div>
      {
        token == '' ?
          <div className='d-flex align-items-center justify-content-center' style={{ height: '100vh' }}>
            <MicrosoftLogin clientId={clientId} authCallback={authHandler} />
          </div> :
          <>
            <div className="container-fluid">
              <div className="w-100 px-0">
                <div className="d-flex p-3 align-items-center justify-content-center">
                  {/* Left Section */}
                  <div className="col-1 ">
                    <div className='d-flex  align-items-center justify-content-center'>
                      <div className="col-1 d-flex align-items-center justify-content-center">
                        {messages.length > 0 && <div className="form-check">
                          <input className="form-check-input" checked={selectedMessages.length == messages.length && selectedMessages.length > 0} onChange={(e) => e.target.checked ? handleSelectAll() : handleUnselectAll()} type="checkbox" value="" id={`flexCheckDefaultAll`} />
                        </div>}
                      </div>
                    </div>
                  </div>
                  {/* Right Section */}
                  <div className="col-11  d-lg-flex align-items-center justify-content-between">

                    <div className="fw-bold text-primary me-2 my-lg-2 my-3"> Outlook Mails</div>
                    <div className="d-lg-flex align-items-center justify-content-end">
                      <div className="me-2 my-2 d-inline-block position-relative">
                        <input
                          type="email"
                          className="form-control searchInput mx-auto me-2"
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
                        <button disabled={searchSubject === ''} onClick={() => getMessages(token)} className='btn btn-primary me-2'>
                          Search
                        </button>
                      </div>
                      <div className="d-inline-block my-2 ">
                        <button onClick={() => getMessages(token)} className='btn btn-primary me-2'>
                          <i className="fa fa-refresh me-2" aria-hidden="true"></i>
                          Refresh
                        </button>
                      </div>
                      {selectedMessages.length > 0 && <div className="d-inline-block my-2 ">
                        <button onClick={() => moveToTrash(token)} className='btn btn-danger me-2'>
                          <i className='fa fa-trash me-2'></i> Trash
                        </button>
                      </div>}
                      <div className="d-inline-block my-2 ">
                        <button onClick={logout} className='btn btn-danger me-2'>
                          Logout
                        </button>
                      </div>

                      <div className="d-inline-block my-2 ">
                        <select className="form-select me-2 mb-0" onChange={(e) => { setTimeFilter(e.target.value) }} id="floatingSelect" aria-label="Floating label select example">
                          <option value="">Filter</option>
                          <option value="24hours">24 Hours</option>
                          <option value="48hours">48 Hours</option>
                          <option value="1week">1 Week</option>
                          <option value="2week">2 Weeks</option>
                          <option value="1month">1 Month</option>
                        </select>
                      </div>
                    </div>
                  </div>

                </div>

              </div>

              {/* Messages Section */}
              <div className="w-100 px-0 bg-light shadow" style={{ minHeight: '100vh' }}>
                {messages.length > 0 ? messages.map((el, i) => (
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
                )) : <div className="message-item border-bottom text-start p-3">
                  <div className="d-flex align-items-center justify-content-center">
                    <div className="col-12">
                      <div className="text-sm row">
                        <div className="col-12 text-center">No Emails Found</div>
                      </div>
                    </div>
                  </div>
                </div>}
              </div>
            </div>


          </>
      }

    </div>
  );
};

export default OutlookMail;
