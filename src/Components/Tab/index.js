import './index.scss'

function Tab({ list, active, setActive }) {
  return (
    <section id="tab">
      <ul>
        {Object.entries(list).map(([key, values]) => (
          <li key={key} className={active === key ? 'active' : ''}>
            <button onClick={() => setActive(key)}>{values.name}</button>
          </li>
        ))}
      </ul>
    </section>
  )
}

export default Tab
