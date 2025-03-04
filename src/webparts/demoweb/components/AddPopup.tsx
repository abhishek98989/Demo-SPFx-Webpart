import React, { useEffect } from 'react'
import { Panel } from '@fluentui/react/lib/Panel';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
export const AddPopup = (props: any) => {
  const sp = spfi().using(SPFx(props.Context));
  const [popupVal, setPopVal]: any = React.useState({
    Title: '',
    Place: '',
    Thing: '',
    Animal: ''
  });
  const saveData = async (popupVal: any) => {
    // Fetch list items
    const items: any = await sp.web.lists.getById('ccc4bcb8-3a9d-4031-bc70-4c267b1ef3ca').items.add({
      Title: popupVal.Title,
      Place: popupVal.Place,
      Thing: popupVal.Thing,
      Animal: popupVal.Animal
    });
    popupVal.Id = items.Id;
    props.addFunc(popupVal);
    clear();
  }
  const DeleteData = async (popupVal: any) => {
    // Fetch list items
    const items: any = await sp.web.lists.getById('ccc4bcb8-3a9d-4031-bc70-4c267b1ef3ca').items.getById(props?.item?.Id).recycle();
    props.DeleteCallback();
  }
  const EditData = async (popupVal: any) => {
    // Fetch list items
    const items: any = await sp.web.lists.getById('ccc4bcb8-3a9d-4031-bc70-4c267b1ef3ca').items.getById(props?.item?.Id).update({
      Title: popupVal.Title,
      Place: popupVal.Place,
      Thing: popupVal.Thing,
      Animal: popupVal.Animal
    });
    props.editCallback(popupVal);
    clear();
  }
  const cancel = () => {
    props.closeFunc();
    clear();
  }
  const clear = () => {
    setPopVal({
      Title: '',
      Place: '',
      Thing: '',
      Animal: ''
    })
  }
  useEffect(() => {
    if (props?.item?.Id != undefined) {
      const fetchItem = async () => {
        const items: any = await sp.web.lists.getById('ccc4bcb8-3a9d-4031-bc70-4c267b1ef3ca').items.getById(props?.item?.Id)();
        setPopVal(items);
      };
      fetchItem();
    }
  }, [props?.item])

  const onRenderFooterContent = React.useCallback(
    (popupVal) => (
      <div>
       {props?.item?.Id != null && <PrimaryButton onClick={() => DeleteData(popupVal)}>Delete</PrimaryButton>}
        {props?.item?.Id == undefined ? <PrimaryButton onClick={() => saveData(popupVal)} >
          Save
        </PrimaryButton> : <PrimaryButton onClick={() => EditData(popupVal)} >Edit</PrimaryButton>}
        <DefaultButton onClick={() => cancel()}>Cancel</DefaultButton>
      </div>
    ),
    [],
  );
  return (
    <Panel
      isOpen={true}
      headerText="Add Data Popup"
      closeButtonAriaLabel="Close"
      onRenderFooterContent={() => onRenderFooterContent(popupVal)}
      // Stretch panel content to fill the available height so the footer is positioned
      // at the bottom of the page
      isFooterAtBottom={true}
    >
      <input placeholder='Enter Name' value={popupVal.Title} onChange={(e) => setPopVal({ ...popupVal, Title: e?.target?.value })} type="text" />
      <input placeholder='Enter Place' value={popupVal.Place} onChange={(e) => setPopVal({ ...popupVal, Place: e?.target?.value })} type="text" />
      <input placeholder='Enter Thing' value={popupVal.Thing} onChange={(e) => setPopVal({ ...popupVal, Thing: e?.target?.value })} type="text" />
      <input placeholder='Enter Animal' value={popupVal.Animal} onChange={(e) => setPopVal({ ...popupVal, Animal: e?.target?.value })} type="text" />
    </Panel>
  )
}
