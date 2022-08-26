import * as React from 'react';
import { useState, useEffect } from "react";
import { Breadcrumb, BreadcrumbDivider, BreadcrumbLink, TeamsIcon } from '@fluentui/react-northstar';
import { IFolder } from "../../../model/IFolder";
import { Folder } from "./Folder";

export const Teams = (props) => {
  const [teams, setTeams] = useState<IFolder[]>();
  const [folders, setFolders] = useState<IFolder[]|null>(null);

  const getJoinedTeams = React.useCallback(async () => {
    props.getJoinedTeams().then((result) => {
      setTeams(result);
      setFolders(null);
    });
  },[props.getJoinedTeams]); // eslint-disable-line react-hooks/exhaustive-deps

  const getFolders = React.useCallback((driveId: string, folderId: string, name: string) => {    
    props.getFolders(driveId, folderId, name).then((result) => {
      setFolders(result);
    });
  },[props.getFolders]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    getJoinedTeams();
  }, []);

  return (
    <div>
      <Breadcrumb>
        <Breadcrumb.Item>
          <BreadcrumbLink className='iconLogo' onClick={() => getJoinedTeams()} ><TeamsIcon /></BreadcrumbLink>
          {props.currentFolder !== null && props.currentFolder.parentFolder !== null &&
              <BreadcrumbDivider />}
          {props.currentFolder !== null && props.currentFolder.parentFolder !== null &&
              <BreadcrumbLink className='breadcrumbFolder' 
                              onClick={() => getFolders(props.currentFolder.parentFolder.driveID, 
                                                        props.currentFolder.parentFolder.id,
                                                        props.currentFolder.parentFolder.name)} >{props.currentFolder.parentFolder.name}</BreadcrumbLink>}
          {props.currentFolder !== null && 
              <BreadcrumbDivider />}
          {props.currentFolder !== null && 
              <BreadcrumbLink onClick={() => getFolders(props.currentFolder.driveID, 
                                                        props.currentFolder.id,
                                                        props.currentFolder.name)} >{props.currentFolder.name}</BreadcrumbLink>}
        </Breadcrumb.Item>
      </Breadcrumb>

      <ul>
        {folders === null && teams?.map(t => {
          return <Folder folder={t} getFolders={getFolders} />
        })}
        {folders !== null && folders?.map(f => {
          return <Folder folder={f} getFolders={getFolders} />
        })}
      </ul>
    </div>
  );
};