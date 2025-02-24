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
  const saveData = () => {
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
    () => (
      <div>
        <PrimaryButton onClick={saveData} >
          Save
        </PrimaryButton>
        <DefaultButton onClick={cancel}>Cancel</DefaultButton>
      </div>
    ),
    [],
  );
  const update = (e:any,poperty:any) => {
    let SaveVal :any= popupVal
    SaveVal[poperty]=e?.target?.value;
    setPopVal(SaveVal);
  }
  return (
    <Panel
      isOpen={true}
      headerText="Add Data Popup"
      closeButtonAriaLabel="Close"
      onRenderFooterContent={onRenderFooterContent}
      // Stretch panel content to fill the available height so the footer is positioned
      // at the bottom of the page
      isFooterAtBottom={true}
    >
      <input placeholder='Enter Name' value={popupVal.name} onChange={(e) => update( e,'name')} type="text" />
      <input placeholder='Enter Place' value={popupVal.place} onChange={(e) =>  update( e,'place')} type="text" />
      <input placeholder='Enter Thing' value={popupVal.thing} onChange={(e) => update( e,'thing')} type="text" />
      <input placeholder='Enter Animal' value={popupVal.animal} onChange={(e) =>  update( e,'animal')} type="text" />
    </Panel>
  )
}
