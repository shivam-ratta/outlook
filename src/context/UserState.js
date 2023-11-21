import React, { useState } from "react";

import UserContext from "./UserContext"

const UserState = (props)=>{
  const [token, setToken] = useState('')
    const update =(data) => {
        setToken(data)
        console.log('updateContext :-', data)
    }
return(
    <UserContext.Provider value={{token,update}}>
        {props.children}
    </UserContext.Provider>
)
}

export default UserState;