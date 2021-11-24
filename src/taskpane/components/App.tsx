import * as React from "react";
import Progress from "./Progress";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

interface PowerPointElement {
  context: PowerPoint.RequestContext;
  load(): void;
}

interface PowerPointCollection<T> extends PowerPointElement {
  items: T[];
}

async function load(item: PowerPointElement) {
  item.load();
  await item.context.sync(); 
}

async function iterate<T>(collection: PowerPointCollection<T>, callback: (item: T) => Promise<void>) {
  await load(collection);
  for (const item of collection.items) {
    await callback(item);
  }
}

const recolor = (originalColor: string, newColor: string) => PowerPoint.run((context) => 
  iterate(context.presentation.slides, (slide) =>
    iterate(slide.shapes, async (shape) => {
      if (shape.textFrame === undefined) {
        throw new Error('This version of PowerPoint does not support the preview API features required to use this add-in');
      }
      try {
        await load(shape);
        await load(shape.textFrame.textRange);
        for (let i = 0; i < shape.textFrame.textRange.text.length; i++) {
          const characterRange = shape.textFrame.textRange.getSubstring(i, 1);
          await load(characterRange.font);
          if (characterRange.font.color.toLowerCase() === originalColor) {
            characterRange.font.color = newColor;
          }
        }
      } catch(e) {}
    })
  )
);

const App = ({ title, isOfficeInitialized }: AppProps) => {
  const [isRecolorRunning, setRecolorRunning] = React.useState(false);
  const [originalColor, setOriginalColor] = React.useState('#ffff00');
  const [newColor, setNewColor] = React.useState('#112fd9');
  const [errorMessage, setErrorMessage] = React.useState<string | null>(null);

  if (!isOfficeInitialized) {
    return (
      <Progress
        title={title}
        logo={require("./../../../assets/logo-filled.png")}
        message="Please sideload your addin to see app body."
      />
    )
  }
  
  const onClick = async () => {
    setRecolorRunning(true);
    try {
      await recolor(originalColor.toLowerCase(), newColor.toLowerCase());
      setErrorMessage(null);
    } catch(e) {
      setErrorMessage(e.message);
    }
    setRecolorRunning(false);
  }
  
  return (
    <main>
      <h1>{title}</h1>
      <label>
        <input
          type="color"
          value={originalColor}
          onChange={(e) => setOriginalColor(e.target.value)}
        />
        Original color
      </label>
      <label>
      <input
          type="color"
          value={newColor}
          onChange={(e) => setNewColor(e.target.value)}
        />
        New color
      </label>
      <button onClick={onClick} disabled={isRecolorRunning}>
        {isRecolorRunning ? 'Recoloring...' : 'Recolor'}
      </button>
      {
        errorMessage && (
          <p className="error">{errorMessage}</p>
        )
      }
    </main>
  )
}

export default App;
