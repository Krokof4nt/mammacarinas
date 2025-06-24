import React, { useState, useEffect } from 'react';
import { Bed, Bath, Soup, Wind, ChevronLeft, ChevronRight, RotateCcw, CheckSquare, Square, Mail } from 'lucide-react';

// --- Configuration ---
const roomNames = [
  "Nils rum", "Görans rum", "Ingvars rum", "Bertils rum", 
  "Ottilias rum", "Kerstins rum", "Charlottes rum", 
  "Camillas rum", "Tant Ruths rum", "Discorummet"
];

// Helper to create a clean key for storing data
const toKey = (name) => name.toLowerCase().replace(/ /g, '-');

const roomChecklistTemplate = {
  title: 'Rummen (Avresestäd)',
  icon: Bed,
  sections: [
    {
      title: 'Säng och Textilier',
      items: [
        { id: 'lakan', text: 'Byt till rena och hela lakan (släta).' },
        { id: 'enhetliga', text: 'Använd enhetliga lakan i samma rum.' },
        { id: 'badda', text: 'Bädda med 2 kuddar per person (1/barn).' },
        { id: 'handdukar', text: 'Lägg fram 1 stor & 1 liten handduk per person.' },
        { id: 'badrumsmatta', text: 'Byt badrumsmatta.' },
      ],
    },
    {
      title: 'Badrum',
      items: [
        { id: 'rengorBadrum', text: 'Rengör toalett, handfat och dusch.' },
        { id: 'tval', text: 'Fyll på med tvål.' },
        { id: 'papper', text: 'Fyll på toalettpapper (minst 4 rullar).' },
        { id: 'glas', text: 'Ställ fram ett rent dricksglas per gäst.' },
      ],
    },
    {
      title: 'Städning & Slutkoll',
      items: [
        { id: 'vadra', text: 'Vädra rummet under städningen.' },
        { id: 'dammsug', text: 'Dammsug & våttorka golvet (även under sängar).' },
        { id: 'torkaYtor', text: 'Torka av alla ytor (lister, dörrar, TV etc).' },
        { id: 'fonster', text: 'Rengör fönster och speglar om kladdiga.' },
        { id: 'sopkorg', text: 'Töm sopkorgen och lägg i ny påse.' },
        { id: 'kvarglomt', text: 'Plocka bort kvarglömda saker.' },
        { id: 'lampor', text: 'Kontrollera att alla lampor fungerar.' },
        { id: 'stangFonster', text: 'Stäng fönstret.' },
        { id: 'slutkoll', text: 'Gör en sista kontroll av rummet.' },
      ],
    },
  ],
};

const otherChecklistsData = {
  allmanToa: {
    title: 'Allmänna Toaletten',
    icon: Bath,
    sections: [{ title: 'Dagligen', items: [
        { id: 'rengorToa', text: 'Rengör toalett och handfat.' },
        { id: 'bytHandduk', text: 'Byt till en ren handduk.' },
        { id: 'fyllPaPapper', text: 'Fyll på toalettpapper (ny rulle i hållaren).' },
    ]}]},
  koket: {
    title: 'Köket',
    icon: Soup,
    sections: [{ title: 'Daglig Rutin', items: [
        { id: 'torkaBord', text: 'Torka av bord och bänkar.' },
        { id: 'bytDisktrasa', text: 'Byt disktrasa.' },
        { id: 'renaHanddukar', text: 'Häng fram rena handdukar.' },
        { id: 'laddaKaffe', text: 'Ladda kaffebryggaren.' },
        { id: 'vattnaBlommor', text: 'Vattna blommorna (vid behov).' },
        { id: 'rensaKyl', text: 'Rensa kylen från gamla varor.' },
    ]}]},
  altan: {
    title: 'Altan',
    icon: Wind,
    sections: [{ title: 'Underhåll', items: [
        { id: 'vattnaAltan', text: 'Vattna blommor.' },
        { id: 'torkaBordAltan', text: 'Torka av bord.' },
        { id: 'vikFiltar', text: 'Skaka och vik filtarna.' },
        { id: 'sopaAltan', text: 'Sopa golvet vid behov.' },
        { id: 'tomSkrap', text: 'Töm askkoppar och skräp.' },
    ]}]},
};

// --- Helper function to initialize state from localStorage or create a default ---
const initializeState = () => {
  const initialState = {};
  
  roomNames.forEach(name => {
    const roomKey = toKey(name);
    initialState[roomKey] = {};
    roomChecklistTemplate.sections.forEach(section => {
      section.items.forEach(item => {
        initialState[roomKey][item.id] = false;
      });
    });
  });

  for (const key in otherChecklistsData) {
    initialState[key] = {};
    otherChecklistsData[key].sections.forEach(section => {
      section.items.forEach(item => {
        initialState[key][item.id] = false;
      });
    });
  }
  return initialState;
};

// --- Components ---

const Modal = ({ isOpen, onClose, onConfirm, title, children }) => {
    if (!isOpen) return null;

    return (
        <div className="fixed inset-0 bg-black bg-opacity-50 z-50 flex justify-center items-center p-4">
            <div className="bg-white dark:bg-gray-800 rounded-xl shadow-xl p-6 w-full max-w-sm">
                <h3 className="text-lg font-bold text-gray-900 dark:text-white mb-4">{title}</h3>
                <div className="text-gray-700 dark:text-gray-300 mb-6">{children}</div>
                <div className="flex justify-end space-x-4">
                    <button onClick={onClose} className="px-4 py-2 rounded-lg text-gray-700 dark:text-gray-200 bg-gray-200 dark:bg-gray-600 hover:bg-gray-300 dark:hover:bg-gray-500 transition-colors">Avbryt</button>
                    <button onClick={onConfirm} className="px-4 py-2 rounded-lg text-white bg-blue-600 hover:bg-blue-700 transition-colors">Bekräfta</button>
                </div>
            </div>
        </div>
    );
};


const Header = ({ title, onBack }) => (
  <header className="bg-white dark:bg-gray-800 shadow-md sticky top-0 z-20">
    <div className="max-w-4xl mx-auto p-2 flex items-center justify-between h-16">
      <div className="flex-1 flex items-center">
        {onBack && (
          <button onClick={onBack} className="p-2 rounded-full hover:bg-gray-200 dark:hover:bg-gray-700">
            <ChevronLeft className="h-6 w-6 text-gray-700 dark:text-gray-300" />
          </button>
        )}
        {title && <h1 className="ml-2 text-lg sm:text-xl font-bold text-gray-900 dark:text-white truncate">{title}</h1>}
      </div>
      
      <div className="flex-shrink-0">
        <img 
            src="https://www.mammacarinas.se/wp-content/uploads/2025/06/Logga-Anpassat.png" 
            alt="Mamma Carinas Logotyp" 
            className="h-12 w-auto"
            onError={(e) => { e.target.onerror = null; e.target.style.display = 'none'; }}
        />
      </div>

      <div className="flex-1"></div>
    </div>
  </header>
);

const MainMenu = ({ onNavigate, onSendAndReset }) => (
    <div className="p-4 sm:p-6 pt-6">
      <div className="flex flex-col gap-3">
        <button onClick={() => onNavigate('room-selection')} className="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-md hover:shadow-lg transition-all transform hover:scale-102 flex items-center justify-center text-center h-20 border border-gray-200 dark:border-gray-700">
            <Bed className="h-7 w-7 mr-4 text-blue-600 dark:text-blue-400" />
            <span className="font-semibold text-lg text-gray-800 dark:text-gray-200">Rummen</span>
        </button>
        {Object.entries(otherChecklistsData).map(([key, { title, icon: Icon }]) => (
          <button key={key} onClick={() => onNavigate(key)} className="bg-white dark:bg-gray-800 p-4 rounded-xl shadow-md hover:shadow-lg transition-all transform hover:scale-102 flex items-center justify-center text-center h-20 border border-gray-200 dark:border-gray-700">
            <Icon className="h-7 w-7 mr-4 text-blue-600 dark:text-blue-400" />
            <span className="font-semibold text-lg text-gray-800 dark:text-gray-200">{title}</span>
          </button>
        ))}
        <button onClick={onSendAndReset} className="mt-4 bg-blue-600 text-white p-4 rounded-xl shadow-lg hover:shadow-xl hover:bg-blue-700 transition-all transform hover:scale-102 flex items-center justify-center text-center h-20">
            <Mail className="h-7 w-7 mr-4" />
            <span className="font-semibold text-lg">Skicka & Nollställ</span>
        </button>
      </div>
    </div>
  );

const RoomSelectionMenu = ({ onNavigate }) => (
    <div className="p-4 sm:p-6 space-y-4">
        {roomNames.map(name => (
            <button
                key={toKey(name)}
                onClick={() => onNavigate(toKey(name))}
                className="w-full bg-white dark:bg-gray-800 p-5 rounded-xl shadow-md hover:shadow-lg transition-all transform hover:scale-102 text-left flex justify-between items-center border border-gray-200 dark:border-gray-700"
            >
                <span className="font-semibold text-gray-800 dark:text-gray-200">{name}</span>
                <ChevronRight className="h-5 w-5 text-gray-400 dark:text-gray-500" />
            </button>
        ))}
    </div>
);

const ChecklistItem = ({ item, isChecked, onToggle }) => (
    <div onClick={onToggle} className="flex items-center p-4 bg-white dark:bg-gray-800 rounded-xl cursor-pointer hover:bg-gray-50 dark:hover:bg-gray-700 transition-colors border border-gray-200 dark:border-gray-700">
        {isChecked ? <CheckSquare className="h-6 w-6 text-blue-600 dark:text-blue-500 flex-shrink-0" /> : <Square className="h-6 w-6 text-gray-400 dark:text-gray-500 flex-shrink-0" />}
        <span className={`ml-4 text-gray-800 dark:text-gray-200 ${isChecked ? 'line-through text-gray-500 dark:text-gray-400' : ''}`}>
            {item.text}
        </span>
    </div>
);

const ChecklistView = ({ checklistKey, checkedItems, onToggle, onReset }) => {
  const isRoom = roomNames.map(toKey).includes(checklistKey);
  const checklistData = isRoom ? roomChecklistTemplate : otherChecklistsData[checklistKey];
  
  return (
    <div className="pb-8">
      {checklistData.sections.map(section => (
        <div key={section.title} className="p-4 sm:p-6">
          <div className="flex justify-between items-center mb-4">
            <h2 className="text-xl font-bold text-gray-900 dark:text-white">{section.title}</h2>
             <button onClick={() => onReset(checklistKey, section.items.map(i => i.id))} className="flex items-center text-xs bg-gray-200 dark:bg-gray-700 hover:bg-gray-300 dark:hover:bg-gray-600 text-gray-700 dark:text-gray-300 font-medium py-1 px-3 rounded-full transition-colors">
                <RotateCcw className="h-3 w-3 mr-1.5" />
                Återställ
            </button>
          </div>
          <div className="space-y-3">
            {section.items.map(item => (
              <ChecklistItem key={item.id} item={item} isChecked={checkedItems[item.id]} onToggle={() => onToggle(checklistKey, item.id)} />
            ))}
          </div>
        </div>
      ))}
    </div>
  );
};

export default function App() {
  const [history, setHistory] = useState(['main']);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [checkedState, setCheckedState] = useState(() => {
    try {
      const storedState = localStorage.getItem('stadrutinerState');
      return storedState ? JSON.parse(storedState) : initializeState();
    } catch (error) {
      console.error("Could not parse localStorage state:", error);
      return initializeState();
    }
  });

  useEffect(() => {
    try {
      localStorage.setItem('stadrutinerState', JSON.stringify(checkedState));
    } catch (error) {
      console.error("Could not save state to localStorage:", error);
    }
  }, [checkedState]);

  const handleNavigate = (view) => setHistory(prev => [...prev, view]);
  const handleBack = () => setHistory(prev => prev.slice(0, -1));
  
  const handleToggle = (listKey, itemId) => {
      setCheckedState(prevState => ({
          ...prevState,
          [listKey]: { ...prevState[listKey], [itemId]: !prevState[listKey][itemId] }
      }));
  };

  const handleReset = (listKey, itemIdsToReset) => {
      setCheckedState(prevState => {
          const newListState = { ...prevState[listKey] };
          itemIdsToReset.forEach(id => { newListState[id] = false; });
          return { ...prevState, [listKey]: newListState };
      });
  };

  const handleSendAndReset = () => {
    let summary = "Hej,\n\nHär är en sammanfattning av dagens slutförda städuppgifter:\n\n";
    let tasksFound = false;

    // Process rooms
    roomNames.forEach(name => {
      const roomKey = toKey(name);
      const roomState = checkedState[roomKey];
      if (!roomState) return;

      const checkedItemsList = [];
      const uncheckedItemsList = [];

      roomChecklistTemplate.sections.forEach(section => {
        section.items.forEach(item => {
          if (roomState[item.id]) {
            checkedItemsList.push(`- ${item.text}`);
          } else {
            uncheckedItemsList.push(`- ${item.text}`);
          }
        });
      });

      if (checkedItemsList.length > 0) {
        tasksFound = true;
        if (uncheckedItemsList.length === 0) {
          summary += `*${name}* - KLART.\n\n`;
        } else {
          summary += `*${name}*\n`;
          summary += `KLART:\n${checkedItemsList.join('\n')}\n\n`;
          summary += `KVARSTÅR:\n${uncheckedItemsList.join('\n')}\n\n`;
        }
      }
    });

    // Process other areas
    Object.entries(otherChecklistsData).forEach(([key, data]) => {
      const areaState = checkedState[key];
      if (!areaState) return;
      
      const checkedItemsList = [];
      const uncheckedItemsList = [];

      data.sections.forEach(section => {
        section.items.forEach(item => {
          if (areaState[item.id]) {
            checkedItemsList.push(`- ${item.text}`);
          } else {
            uncheckedItemsList.push(`- ${item.text}`);
          }
        });
      });
      
      if (checkedItemsList.length > 0) {
          tasksFound = true;
          if (uncheckedItemsList.length === 0) {
            summary += `*${data.title}* - KLART.\n\n`;
          } else {
            summary += `*${data.title}*\n`;
            summary += `KLART:\n${checkedItemsList.join('\n')}\n\n`;
            summary += `KVARSTÅR:\n${uncheckedItemsList.join('\n')}\n\n`;
          }
      }
    });
    
    if (!tasksFound) {
      summary = "Inga uppgifter har markerats som slutförda idag.";
    }

    const email = "info@mammarinas.se";
    const subject = `Städrapport - ${new Date().toLocaleDateString('sv-SE')}`;
    const mailtoLink = `mailto:${email}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(summary)}`;

    window.location.href = mailtoLink;

    setTimeout(() => {
        setIsModalOpen(true);
    }, 500);
  };
  
  const handleConfirmResetAll = () => {
    setCheckedState(initializeState());
    setIsModalOpen(false);
  };

  const currentView = history[history.length - 1];

  const renderContent = () => {
    if (currentView === 'main') return <MainMenu onNavigate={handleNavigate} onSendAndReset={handleSendAndReset} />;
    if (currentView === 'room-selection') return <RoomSelectionMenu onNavigate={handleNavigate} />;
    if (checkedState[currentView]) {
      return <ChecklistView checklistKey={currentView} checkedItems={checkedState[currentView]} onToggle={handleToggle} onReset={handleReset} />;
    }
    return <MainMenu onNavigate={handleNavigate} onSendAndReset={handleSendAndReset} />; // Fallback
  };
  
  const getTitle = () => {
    if (currentView === 'room-selection') return 'Välj rum';
    const isRoom = roomNames.map(toKey).includes(currentView);
    if(isRoom) return roomNames.find(name => toKey(name) === currentView);
    return otherChecklistsData[currentView]?.title || 'Städrutiner';
  }

  return (
    <main className="bg-gray-100 dark:bg-gray-900 min-h-screen font-sans">
      <div className="max-w-4xl mx-auto bg-gray-100 dark:bg-gray-900">
        <Header 
            title={currentView !== 'main' ? getTitle() : null}
            onBack={currentView !== 'main' ? handleBack : null}
        />
        {renderContent()}
        <Modal 
            isOpen={isModalOpen} 
            onClose={() => setIsModalOpen(false)} 
            onConfirm={handleConfirmResetAll}
            title="Nollställ alla checklistor?"
        >
            <p>Vill du nollställa alla ibockade rutor i hela appen? Detta rensar alla checklistor.</p>
        </Modal>
      </div>
    </main>
  );
}
