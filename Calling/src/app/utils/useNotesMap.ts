import { useEffect, useState, useCallback } from 'react';
import { debounce } from 'lodash';
import { SharedMap } from 'fluid-framework';

export const useNotesMap = (notesMap?: SharedMap) => {
  const [notes, setNotes] = useState<any[]>([]);
  const [started, setStarted] = useState(false);
  const addNote = useCallback(
    (id: string, note: any) => {
      if (!notes.find((item) => item.id === id)) {
        notesMap?.set(id, note);
      }
    },
    [notesMap]
  );

  const removeNote = useCallback(
    (id: string) => {
      notesMap?.delete(id);
    },
    [notesMap]
  );

  const refreshNotes = useCallback(() => {
    const items: any[] = [];
    notesMap?.forEach((value, key) => {
      items.push(value);
    });
    setNotes(items);
  }, [notesMap, setNotes]);

  // eslint-disable-next-line
  const debouncedRefresh = useCallback(debounce(refreshNotes, 100), [refreshNotes]);

  useEffect(() => {
    if (notesMap && !started) {
      notesMap.on('valueChanged', debouncedRefresh);
      debouncedRefresh();
      console.log('useNotesMap: started notes');
      setStarted(true);
    }
  }, [notesMap, started, setStarted, debouncedRefresh]);

  return {
    started,
    notes,
    addNote,
    removeNote
  };
};
