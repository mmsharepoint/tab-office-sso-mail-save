import * as React from "react";
import { Provider, Flex, Text, Button, Header, List, ListItemProps, PaperclipIcon, Dialog } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import { app, authentication } from "@microsoft/teams-js";
import Axios from "axios";
import { IMail } from "../../model/IMail";
import { OneDrive } from "./components/OneDrive";
import { IFolder } from "../../model/IFolder";

/**
 * Implementation of the Mail Storage Tab content page
 */
export const MailStorageTab = () => {

  const [{ inTeams, theme, context }] = useTeams();
  const [entityId, setEntityId] = useState<string | undefined>();
  const [token, setToken] = useState<string>();
  const [error, setError] = useState<string>();
  const [mails, setMails] = React.useState<IMail[]>([]);
  const [mailItems, setMailItems] = React.useState<ListItemProps[]>([]);
  const [selectedIndex, setSelectedIndex] = React.useState<number>();
  const [currentFolder, setCurrentFolder] = React.useState<IFolder|null>(null);

  const attachmentIcon = <PaperclipIcon />;

  const getMails = async (token: string) => {
    const response = await Axios.get(`https://${process.env.PUBLIC_HOSTNAME}/api/mails`,
    { headers: { Authorization: `Bearer ${token}` }});

    setMails(response.data);
  };

  const saveMail = async () => {
    alert(
      `Mail to save has id "${mails[selectedIndex!].id}" and subject "${mails[selectedIndex!].subject}"`,
    );
    let requestUrl = `https://${process.env.PUBLIC_HOSTNAME}/api/mail/${mails[selectedIndex!].id}`;
    if (currentFolder === null) {
      requestUrl += "/*/*"
    }
    else {
      requestUrl += `/${currentFolder.driveID}/${currentFolder.id}`
    }
    const response = await Axios.post(requestUrl, {},
    { headers: { Authorization: `Bearer ${token}` }});

  };

  const getFolders = async (driveId: string, folderId: string, name: string) => {
    const response = await Axios.get(`https://${process.env.PUBLIC_HOSTNAME}/api/folders/${driveId}/${folderId}`,
    { headers: { Authorization: `Bearer ${token}` }});
    if (driveId !== "*" && folderId !== "*") {
      setCurrentFolder({id: folderId, driveID: driveId, parentFolder: currentFolder, name: name})
    }
    return response.data;
  };

  useEffect(() => {
    if (inTeams === true) {
      authentication.getAuthToken({
          resources: [`api://${process.env.PUBLIC_HOSTNAME}/${process.env.TAB_APP_ID}`],
          silent: false
      } as authentication.AuthTokenRequestParameters).then(token => {
        getMails(token);
        setToken(token);
        app.notifySuccess();
      }).catch(message => {
          setError(message);
          app.notifyFailure({
              reason: app.FailedReason.AuthFailed,
              message
          });
      });
    } else {
        setEntityId("Not in Microsoft Teams");
    }
  }, [inTeams]);

  useEffect(() => {
    if (context) {
      setEntityId(context.page.id);
    }
  }, [context]);

  useEffect(() => {
    if (mails.length > 0) {
      let listItems: ListItemProps[] = [];
      mails.forEach((m) => {
        listItems.push({ header: m.from, content: m.subject, media: m.hasAttachments ? (attachmentIcon) : "", headerMedia: m.receivedDateTime });
      });
      setMailItems(listItems);
    }
  }, [mails]);

  /**
   * The render() method to create the UI of the tab
   */
  return (
    <Provider theme={theme}>
      <Flex fill={true} column styles={{
          padding: ".8rem 0 .8rem .5rem"
      }}>
        <Flex.Item>
        <Dialog
            cancelButton="Cancel"
            confirmButton="Save here"
            content={<OneDrive currentFolder={currentFolder} getFolders={getFolders} />}
            // onCancel={onCancel}
            onConfirm={saveMail}
            // onOpen={onOpen}
            // open={open}
            header="Action confirmation"
            trigger={<Button content="Select folder" />}
          />
          
        </Flex.Item>
        <Flex.Item>
          <div>
            <List
              selectable
              selectedIndex={selectedIndex}
              onSelectedIndexChange={(e, newProps) => {                    
                setSelectedIndex(newProps!.selectedIndex);
              }}
              items={mailItems}
            />
            <Button content="Save Mail" primary onClick={saveMail} disabled={typeof selectedIndex !== undefined && (selectedIndex!<0)} />
              {error && <div><Text content={`An SSO error occurred ${error}`} /></div>}
          </div>
        </Flex.Item>
      </Flex>
    </Provider>
  );
};
