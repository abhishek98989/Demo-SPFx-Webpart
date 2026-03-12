import * as React from "react";
import { useState, useEffect } from "react";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { PrimaryButton, IconButton } from "@fluentui/react";
import { ItemPopup } from "./ItemPopup";

type Row = {
  Name: string;
  Location: string;
  Fruit: string;
  Color: string;
};

export const Newtable = (props: any) => {
  const sp = spfi('https://standardcharteredbank.sharepoint.com/sites/fcso_bm').using(SPFx(props?.Context));
  const listID = "7758D77E-9E63-46A7-B660-20AA68C9CA0B";
  // const sp = spfi("https://hhhhteams.sharepoint.com/sites/HHHH/test")
  //   .using(SPFx(props?.Context));

  // const listID = "0CA34243-C724-47FF-AC70-8E79242B9C3C";

  const [rows, setRows] = useState<Row[]>([]);
  const [show, setShow] = useState(false);

  const getdata = async () => {
    const items: Row[] = await sp.web.lists
      .getById(listID)
      .items.select("Name", "Location", "Fruit", "Color", "ID")();

    setRows(items);
  };

  useEffect(() => {
    getdata();
  }, []);

  const handleClose = () => setShow(false);
  const handleShow = () => setShow(true);

  const onSave = async (data: Row) => {
    if (
      data.Name &&
      data.Location &&
      data.Fruit &&
      data.Color
    ) {
      setRows(prev => [...prev, data]);
    }
    setShow(false);
  };
const onDelete = async (row:any) => {
  if(confirm("Are you sure you want to delete this item?")){
    const itemId = row.ID;
    await sp.web.lists.getById(listID).items.getById(itemId).recycle();
    setRows(prev => prev.filter((item:any) => item.ID !== itemId));
  }
}
  return (
    <div>
      <PrimaryButton text="Add New Row" onClick={handleShow} />

      {show && (
        <ItemPopup
          saveCallBack={onSave}
          closeCallback={handleClose}
          Context={props.Context}
        />
      )}

      <table
        style={{
          width: "100%",
          marginTop: "16px",
          borderCollapse: "collapse"
        }}
      >
        <thead>
          <tr>
            <th style={thStyle}>Name</th>
            <th style={thStyle}>Location</th>
            <th style={thStyle}>Fruit</th>
            <th style={thStyle}>Color</th>
            <th style={thStyle}>Action</th>
          </tr>
        </thead>

        <tbody>
          {rows.length > 0 ? (
            rows.map((row, index) => (
              <tr key={index}>
                <td style={tdStyle}>{row.Name}</td>
                <td style={tdStyle}>{row.Location}</td>
                <td style={tdStyle}>{row.Fruit}</td>
                <td style={tdStyle}>{row.Color}</td>
                <td style={tdStyle}>
                  <IconButton
                    iconProps={{ iconName: "Edit" }}
                    title="Edit"
                    ariaLabel="Edit"
                    onClick={() => console.log("Edit row", row)}
                  />

                  <IconButton
                    iconProps={{ iconName: "Delete" }}
                    title="Delete"
                    ariaLabel="Delete"
                    styles={{ root: { color: "red" } }}
                    onClick={() => onDelete(row)}
                  />
                </td>

              </tr>
            ))
          ) : (
            <tr>
              <td colSpan={4} style={{ textAlign: "center", padding: "8px" }}>
                No data available
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
};

// simple styles
const thStyle: React.CSSProperties = {
  border: "1px solid #ddd",
  padding: "8px",
  backgroundColor: "#f3f2f1",
  textAlign: "left"
};

const tdStyle: React.CSSProperties = {
  border: "1px solid #ddd",
  padding: "8px"
};
