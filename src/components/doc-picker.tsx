/** The documents a screen keeps (analyses, decks, PowerPoints): choose, create, rename, delete. */
import { PickerMenu } from "./pickers";

export function DocPicker<T extends { id: string }>({ items, current, name, onSelect, onCreate, onDelete, onRename, createLabel, empty }: { items: T[]; current?: T; name: (t: T) => string; onSelect: (id: string) => void; onCreate: () => void; onDelete: (id: string) => void; onRename: (t: T, name: string) => void; createLabel: string; empty: string }) {
  return <PickerMenu items={items} current={current} name={name} onSelect={onSelect} onRename={onRename} onDelete={onDelete} addLabel={createLabel} onAdd={onCreate} empty={empty} />;
}
