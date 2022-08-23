import * as React from 'react';
import { useState, useEffect } from "react";
import { TabsIcon } from '@fluentui/react-icons-northstar'
import { IFolder } from "../../../model/IFolder";

export const Folder = (props) => {
  const getFolder = () => {
    props.getFolders(props.folder.driveID, props.folder.id, props.folder.name);
  };

  return (
    <li>
      <TabsIcon />                
      <span onClick={getFolder}>{props.folder.name}</span>
    </li>
  );
}