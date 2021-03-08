import * as React from 'react';
import styles from './ReadJsonFile.module.scss';
import { IReadJsonFileProps } from './IReadJsonFileProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DefaultButton, PrimaryButton, IStackTokens } from 'office-ui-fabric-react';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";


export default function ReadJsonFile(props: IReadJsonFileProps) {
  const [filePath, setfilePath] = React.useState('');
  const [fileName, setFileName] = React.useState('');

  const onChangeFilePathHandler = event => {
    setfilePath(event.target.value);
  };
  const onChangeFileNameHandler = event => {
    setFileName(event.target.value);
  };
  return (
    <div className={styles.readJsonFile}>
      <div className={styles.container}>
        <div className={styles.row}>
          <div className={styles.column}>
            <TextField label="Document path" value={filePath} onChange={onChangeFilePathHandler} placeholder="Please enter server relative path of the document" />
          </div><br /><br /><br />
          <div className={styles.column}>
            <TextField label="File Name" value={fileName} onChange={onChangeFileNameHandler} placeholder="Please enter file name" />
          </div><br /><br /><br />
          <div className={styles.column}>
            <PrimaryButton text="Read JSON File" onClick={_readJSONFile} />
          </div>
        </div>
      </div>
    </div>
  );

  async function _readJSONFile() {
    const text2: string = await sp.web.getFolderByServerRelativeUrl(filePath).files.getByName(fileName).getText();
    alert(text2);
  }


}

