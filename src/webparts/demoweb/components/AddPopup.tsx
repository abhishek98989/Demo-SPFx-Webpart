import React from 'react'
import { Panel } from '@fluentui/react/lib/Panel';
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
export const AddPopup = (props: any) => {
  const [popupVal, setPopVal] = React.useState({
    name: '',
    place: '',
    thing: '',
    animal: ''
  });
  const saveData = (popupVal:any) => {
    props.addFunc(popupVal);
    setPopVal({
      name: '',
      place: '',
      thing: '',
      animal: ''
    })
  }
  const cancel = () => {
    props.closeFunc();
    setPopVal({
      name: '',
      place: '',
      thing: '',
      animal: ''
    })
  }
  const onRenderFooterContent = React.useCallback(
    (popupVal) => (
      <div>
        <PrimaryButton onClick={()=>saveData(popupVal)} >
          Save
        </PrimaryButton>
        <DefaultButton onClick={()=>cancel()}>Cancel</DefaultButton>
      </div>
    ),
    [],
  );
  return (
    <Panel
      isOpen={true}
      headerText="Add Data Popup"
      closeButtonAriaLabel="Close"
      onRenderFooterContent={()=>onRenderFooterContent(popupVal)}
      // Stretch panel content to fill the available height so the footer is positioned
      // at the bottom of the page
      isFooterAtBottom={true}
    >
      <input placeholder='Enter Name' value={popupVal.name} onChange={(e) => setPopVal({ ...popupVal, name: e?.target?.value })} type="text" />
      <input placeholder='Enter Place' value={popupVal.place} onChange={(e) => setPopVal({ ...popupVal, place: e?.target?.value })} type="text" />
      <input placeholder='Enter Thing' value={popupVal.thing} onChange={(e) => setPopVal({ ...popupVal, thing: e?.target?.value })} type="text" />
      <input placeholder='Enter Animal' value={popupVal.animal} onChange={(e) => setPopVal({ ...popupVal, animal: e?.target?.value })} type="text" />
    </Panel>
  )
}
