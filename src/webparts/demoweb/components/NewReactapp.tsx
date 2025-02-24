import React from 'react'
import $ from 'jquery'
import { AddPopup } from './AddPopup';
import { set } from '@microsoft/sp-lodash-subset';
import { PrimaryButton } from '@fluentui/react';
export const NewReactapp = () => {
  const [val, setVal] = React.useState(0);
  const [isPopupOpen, setIsPopupOpen] = React.useState(false);
  const [gameValues, setGameValues] = React.useState([{
    name: "Abhishek",
    place: "Noida",
    thing: "Spfx",
    animal: "Animal"
  }]);

  const update = () => {
    setVal(20);
  }
  React.useEffect(() => {
    update();
  }, []);

  const addData = (data: any) => {
    let gameData = gameValues;
    gameData.push(data);
    setGameValues(gameData);
    // setGameValues([...gameValues, data]);
    // setGameValues((prev) => [...prev, data]);
    setIsPopupOpen(false);
  }
  const closePopup = () => {
    setIsPopupOpen(false);
  }

  return (
    <>
      <div className="">
        <PrimaryButton onClick={()=> setIsPopupOpen(true)} >
          Add Data
        </PrimaryButton>
        <table>
          <tr>
            <th>Name</th>
            <th>Place</th>
            <th>Thing</th>
            <th>Animal</th>
          </tr>
          {gameValues.map((item) => {
            return (
              <tr>
                <td>{item.name}</td>
                <td>{item.place}</td>
                <td>{item.thing}</td>
                <td>{item.animal}</td>
              </tr>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      
            )
          })}
        </table>
      </div>
      {isPopupOpen == true && <AddPopup closeFunc={closePopup} addFunc={addData}/>}
    </>
  )
}
