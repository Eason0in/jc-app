import { useState } from 'react'
import './App.scss'
import Beam from './Components/Beam'
import Column from './Components/Column'
import Tab from './Components/Tab'

function App() {
  const [active, setActive] = useState('beam')
  const list = {
    beam: {
      name: '樑',
      elem: <Beam />,
    },
    column: {
      name: '柱',
      elem: <Column />,
    },
    wall: {
      name: '牆',
      elem: <></>,
    },
    board: {
      name: '板',
      elem: <></>,
    },
  }
  return (
    <div className="container">
      <Tab list={list} active={active} setActive={setActive} />
      {list[active].elem}
    </div>
  )
}

export default App
