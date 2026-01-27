import * as React from "react";
import { useState } from "react";
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

type Row = { Name: string; Place: string; Fruit: string; Color: string };

export const Newtable = (props: any) => {
  const [show, setShow] = useState(false);

  const [rows, setRows] = useState<Row[]>([
    { Name: "ABC", Place: "XTZ", Fruit: "MNO", Color: "QRT" },
    { Name: "Senthil", Place: "Chennai", Fruit: "Apple", Color: "Red" }
  ]);

  const [data, setData] = useState<Row>({
    Name: "ABC",
    Place: "XYZ",
    Fruit: "MNO",
    Color: "QRT"
  });

  const handleClose = () => setShow(false);
  const handleShow = () => setShow(true);

  const onChangeField =
    (key: keyof Row) =>
    (_e: any, value?: string) => {
      setData((prev) => ({ ...prev, [key]: value ?? "" }));
    };

  const onSave = () => {
    setRows((prev) => [...prev, { ...data }]);
    setShow(false);
  };

  const columns: IColumn[] = [
    { key: "c1", name: "Name", fieldName: "Name", minWidth: 100, isResizable: true },
    { key: "c2", name: "Place", fieldName: "Place", minWidth: 100, isResizable: true },
    { key: "c3", name: "Fruit", fieldName: "Fruit", minWidth: 100, isResizable: true },
    { key: "c4", name: "Color", fieldName: "Color", minWidth: 100, isResizable: true }
  ];

  return (
    <div>
      <PrimaryButton text="Add New Row" onClick={handleShow} />

      <Dialog
        hidden={!show}
        onDismiss={handleClose}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Add new Data In Table"
        }}
        modalProps={{
          isBlocking: false
        }}
      >
        <Stack tokens={{ childrenGap: 10 }}>
          <TextField label="Name" value={data.Name} onChange={onChangeField("Name")} />
          <TextField label="Place" value={data.Place} onChange={onChangeField("Place")} />
          <TextField label="Fruit" value={data.Fruit} onChange={onChangeField("Fruit")} />
          <TextField label="Color" value={data.Color} onChange={onChangeField("Color")} />
        </Stack>

        <DialogFooter>
          <DefaultButton text="Close" onClick={handleClose} />
          <PrimaryButton text="Save Changes" onClick={onSave} />
        </DialogFooter>
      </Dialog>

      <div style={{ marginTop: 16 }}>
        <DetailsList
          items={rows}
          columns={columns}
          setKey="rows"
          layoutMode={DetailsListLayoutMode.fixedColumns}
        />
      </div>
    </div>
  );
};
