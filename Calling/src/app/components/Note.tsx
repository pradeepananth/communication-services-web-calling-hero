// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/explicit-module-boundary-types */
/* eslint-disable react/prop-types */

import React from 'react';
import '../styles/note.css';

export const Note = ({ index, text, color, onDelete }) => {
  // function to handle delete button click
  const handleClick = () => {
    // call the onDelete prop with the index of the note
    onDelete(index);
  };

  return (
    // render a div with the note class and the given color
    <div className="note" style={{ backgroundColor: color }}>
      {/* render the text prop as the note content */}
      <p>{text}</p>
      {/* render a button with the delete-button class and the handleClick function */}
      <button className="delete-button" onClick={handleClick}></button>
    </div>
  );
};

export default Note;
