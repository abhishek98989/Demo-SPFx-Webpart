import React from 'react'
import $ from 'jquery'
import { AddPopup } from './AddPopup';
import { set } from '@microsoft/sp-lodash-subset';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { PrimaryButton } from '@fluentui/react';
export const NewReactapp = (props: any) => {
  const [val, setVal] = React.useState(0);

  const [isPopupOpen, setIsPopupOpen] = React.useState(false);
  const [gameValues, setGameValues] = React.useState([{
    Title: "Abhishek",
    Place: "Noida",
    Thing: "Spfx",
    Animal: "Animal"
  }]);
  const [editItem, setEditItem]: any = React.useState(undefined);
  const [editIndex, setEditIndex] = React.useState(0);
  const update = () => {
    setVal(20);
  } 
  React.useEffect(() => {
    getItemsfromList();
  }, []);

  const addData = (data: any) => {
    let gameData = gameValues;
    gameData.push(data);
    setGameValues(gameData);
    // getItemsfromList();
    // setGameValues([...gameValues, data]);
    // setGameValues((prev) => [...prev, data]);
    setIsPopupOpen(false);
  }
  const closePopup = () => {
    setIsPopupOpen(false);
  }
  // const editFunc = (item:any,index:number) => {
  const editFunc = (item: any) => {
    setIsPopupOpen(true);
    setEditItem(item);
  }
  const editCallback = (data: any) => {
    let gameData = gameValues;
    // gameData[editIndex] = data;
    gameData = gameData?.map((item: any) => {
      if (item.Id == editItem.Id) {
        return { ...item, ...data };
      } else {
        return item;
      }
    })
    setGameValues(gameData);
    setIsPopupOpen(false);
    setEditItem(undefined);
    // getItemsfromList();
    // setEditIndex(0);
  }
  const DeleteCallback = () => {
    let gameData = gameValues;
    gameData = gameData?.filter((item: any) => item.Id != editItem.Id)
    setGameValues(gameData);
    setIsPopupOpen(false);
    setEditItem(undefined);
    // getItemsfromList();
    // setEditIndex(0);
  }
  const getItemsfromList = async () => {
    // Fetch list items
    const sp = spfi().using(SPFx(props.Context));
    const items: any = await sp.web.lists.getById('ccc4bcb8-3a9d-4031-bc70-4c267b1ef3ca').items.select("Title,Place,Thing,Animal,Id").filter<any>(f => f.text("Animal").equals("Dog").or().text('Place').equals('Noida'))();
    if (items?.length > 0) {
      setGameValues(items);
    }
    console.log(items);
  }
  return (
    <>
      <div className="">
        <PrimaryButton onClick={() => setIsPopupOpen(true)} >
          Add Data
        </PrimaryButton>
        <table>
          <tr>
            <th>Name</th>
            <th>Place</th>
            <th>Thing</th>
            <th>Animal</th>
            <th>Edit</th>
          </tr>
          {gameValues.map((item, index) => {
            return (
              <tr>
                <td>{item.Title}</td>
                <td>{item.Place}</td>
                <td>{item.Thing}</td>
                <td>{item.Animal}</td>
                {/* <td onClick={()=>editFunc(item,index)}>{'=>'}</td> */}
                <td onClick={() => editFunc(item)}>{'=>'}</td>
              </tr>
            )
          })}
        </table>
      </div>
      {isPopupOpen == true && <AddPopup closeFunc={closePopup} editCallback={editCallback} DeleteCallback={DeleteCallback} Context={props?.Context} item={editItem} addFunc={addData} />}
    </>
  )
}
