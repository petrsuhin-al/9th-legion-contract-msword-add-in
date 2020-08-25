import React, { ReactElement, useState } from 'react';
import { DefaultButton, ButtonType } from 'office-ui-fabric-react';

import Header from './Header';
import HeroList from './HeroList';
import Progress from './Progress';

export default function App ({ title, isOfficeInitialized }): ReactElement {
  const [listItems] = useState([
    { icon: 'Ribbon', primaryText: 'Achieve more with Office integration' },
    { icon: 'Unlock', primaryText: 'Unlock features and functionality' },
    { icon: 'Design', primaryText: 'Create and visualize like a pro' }
  ]);

  const handleClick = () => Word.run(async (context) => {
    const paragraph = context.document.body.insertParagraph('Hello World', Word.InsertLocation.end);
    paragraph.font.color = 'blue';
    await context.sync();
  });

  if (!isOfficeInitialized) {
    return (
      <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
    );
  }

  return (
    <div className="ms-welcome">
      <Header logo="assets/logo-filled.png" title={title} message="Welcome" />
      <HeroList message="Discover what Office Add-ins can do for you today!" items={listItems}>
        <p className="ms-font-l">
          Modify the source files, then click <b>Run</b>.
        </p>
        <DefaultButton
          className="ms-welcome__action"
          buttonType={ButtonType.hero}
          iconProps={{ iconName: 'ChevronRight' }}
          onClick={handleClick}
        >
          Run
        </DefaultButton>
      </HeroList>
    </div>
  );
}
