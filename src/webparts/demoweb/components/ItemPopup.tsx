import React, { useState } from 'react'
import {
    PrimaryButton,
    DefaultButton,
    Dialog,
    DialogType,
    DialogFooter,
    TextField,
    DetailsList,
    DetailsListLayoutMode,
    IColumn,
    Stack
} from "@fluentui/react";
import { spfi, SPFx } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
export const ItemPopup = (props: any) => {
      const sp = spfi('https://standardcharteredbank.sharepoint.com/sites/fcso_bm').using(SPFx(props?.Context));
  const listID = "7758D77E-9E63-46A7-B660-20AA68C9CA0B";


    
    // const sp = spfi('https://hhhhteams.sharepoint.com/sites/HHHH/test').using(SPFx(props?.Context));
    //   const listID = "0CA34243-C724-47FF-AC70-8E79242B9C3C";
    const [data, setData]: any = useState({
        Name: "",
        Location: "",
        Fruit: "",
        Color: ""
    });
    
    const onSave = async () => {
        if (
            data.Name.length !== 0 && data.Location.length !== 0 && data.Color.length !== 0 && data.Fruit.length !== 0
        ) {
                await sp.web.lists.getById(listID).items.add({
        //Title: "Test 1",
        Name: data.Name,
        Location: data.Location,
        Fruit: data.Fruit,
        Color: data.Color,
      });
        props?.saveCallBack(data)
            setData({
                Name: "",
                Location: "",
                Fruit: "",
                Color: ""
            })
          
        }

        else {
            alert("Please enter value in all the fields")
        }

    }
    return (
        <Dialog
        hidden={false}
            onDismiss={undefined}
            dialogContentProps={{
                type: DialogType.normal,
                title: "Add new Data In Table"
            }}
            modalProps={{
                isBlocking: false
            }}
        >
            <Stack tokens={{ childrenGap: 10 }}>
                <input type="text" placeholder="Name" value={data.Name} onChange={(e) => setData({ ...data, Name: e.target.value })} />
                <input type="text" placeholder="Location" value={data.Location} onChange={(e) => setData({ ...data, Location: e.target.value })} />
                <input type="text" placeholder="Fruit" value={data.Fruit} onChange={(e) => setData({ ...data, Fruit: e.target.value })} />
                <input type="text" placeholder="Color" value={data.Color} onChange={(e) => setData({ ...data, Color: e.target.value })} />
            </Stack>

            <DialogFooter>
                <DefaultButton text="Close" onClick={props?.closeCallback} />
                <PrimaryButton text="Save Changes" onClick={onSave} />
            </DialogFooter>
        </Dialog>

    )
}
