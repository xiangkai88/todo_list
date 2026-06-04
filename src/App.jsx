import { useEffect, useMemo, useRef, useState } from 'react'

const STORAGE_KEY = 'todo-reminder-items'

const sampleTasks = [
  {
    id: crypto.randomUUID(),
    title: '准备周会纪要',
    note: '整理昨天的进展和今天的阻塞项',
    dueAt: nextLocalTime(1, 0),
    priority: 'high',
    completed: false,
    notified: false,
    createdAt: new Date().toISOString(),
  },
  {
    id: crypto.randomUUID(),
    title: '补充饮水提醒',
    note: '每天下午记得喝水和起身活动',
    dueAt: nextLocalTime(3, 30),
    priority: 'medium',
    completed: false,
    notified: false,
    createdAt: new Date().toISOString(),
  },
]

function nextLocalTime(hoursToAdd, minutesToAdd) {
  const next = new Date()
  next.setHours(next.getHours() + hoursToAdd)
  next.setMinutes(next.getMinutes() + minutesToAdd)
  next.setSeconds(0, 0)
  return next.toISOString()
}

function formatDateTime(value) {
  return new Intl.DateTimeFormat('zh-CN', {
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  }).format(new Date(value))
}

function getMinutesLeft(value) {
  const diff = new Date(value).getTime() - Date.now()
  return Math.round(diff / 60000)
}

function readInitialTasks() {
  const raw = localStorage.getItem(STORAGE_KEY)
  if (!raw) {
    return sampleTasks
  }

  try {
    const parsed = JSON.parse(raw)
    if (Array.isArray(parsed)) {
      return parsed
    }
  } catch {
    return sampleTasks
  }

  return sampleTasks
}

function priorityLabel(priority) {
  if (priority === 'high') return '高优先级'
  if (priority === 'medium') return '中优先级'
  return '低优先级'
}

export default function App() {
  const [tasks, setTasks] = useState(readInitialTasks)
  const [filter, setFilter] = useState('all')
  const [permission, setPermission] = useState(
    typeof Notification === 'undefined' ? 'unsupported' : Notification.permission,
  )
  const [form, setForm] = useState({
    title: '',
    note: '',
    dueAt: '',
    priority: 'medium',
  })
  const intervalRef = useRef(null)

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(tasks))
  }, [tasks])

  useEffect(() => {
    if (typeof Notification === 'undefined') {
      return undefined
    }

    intervalRef.current = window.setInterval(() => {
      setTasks((currentTasks) =>
        currentTasks.map((task) => {
          const isDue = !task.completed && !task.notified && new Date(task.dueAt).getTime() <= Date.now()
          if (!isDue) {
            return task
          }

          new Notification(`待办提醒：${task.title}`, {
            body: task.note || `截止时间：${formatDateTime(task.dueAt)}`,
          })

          return { ...task, notified: true }
        }),
      )
    }, 30000)

    return () => window.clearInterval(intervalRef.current)
  }, [])

  const visibleTasks = useMemo(() => {
    if (filter === 'pending') {
      return tasks.filter((task) => !task.completed)
    }
    if (filter === 'completed') {
      return tasks.filter((task) => task.completed)
    }
    if (filter === 'today') {
      const today = new Date().toDateString()
      return tasks.filter((task) => new Date(task.dueAt).toDateString() === today)
    }
    return tasks
  }, [filter, tasks])

  const stats = useMemo(() => {
    const completed = tasks.filter((task) => task.completed).length
    const overdue = tasks.filter(
      (task) => !task.completed && new Date(task.dueAt).getTime() < Date.now(),
    ).length
    return {
      total: tasks.length,
      completed,
      pending: tasks.length - completed,
      overdue,
    }
  }, [tasks])

  function requestPermission() {
    if (typeof Notification === 'undefined') {
      return
    }

    Notification.requestPermission().then((value) => {
      setPermission(value)
    })
  }

  function handleSubmit(event) {
    event.preventDefault()
    if (!form.title.trim() || !form.dueAt) {
      return
    }

    setTasks((currentTasks) => [
      {
        id: crypto.randomUUID(),
        title: form.title.trim(),
        note: form.note.trim(),
        dueAt: new Date(form.dueAt).toISOString(),
        priority: form.priority,
        completed: false,
        notified: false,
        createdAt: new Date().toISOString(),
      },
      ...currentTasks,
    ])

    setForm({
      title: '',
      note: '',
      dueAt: '',
      priority: 'medium',
    })
  }

  function toggleTask(id) {
    setTasks((currentTasks) =>
      currentTasks.map((task) =>
        task.id === id
          ? {
              ...task,
              completed: !task.completed,
            }
          : task,
      ),
    )
  }

  function removeTask(id) {
    setTasks((currentTasks) => currentTasks.filter((task) => task.id !== id))
  }

  function clearCompleted() {
    setTasks((currentTasks) => currentTasks.filter((task) => !task.completed))
  }

  return (
    <div className="page-shell">
      <main className="app-card">
        <section className="hero-panel">
          <div>
            <p className="eyebrow">Todo Reminder</p>
            <h1>网页版代办事项提醒工具</h1>
            <p className="hero-copy">
              记录任务、设置截止时间，并在浏览器里接收提醒。所有数据保存在本地浏览器中。
            </p>
          </div>
          <button
            className="notification-button"
            type="button"
            onClick={requestPermission}
            disabled={permission === 'granted' || permission === 'unsupported'}
          >
            {permission === 'granted'
              ? '提醒权限已开启'
              : permission === 'denied'
                ? '提醒权限被拒绝'
                : permission === 'unsupported'
                  ? '当前浏览器不支持提醒'
                  : '开启浏览器提醒'}
          </button>
        </section>

        <section className="stats-grid">
          <article>
            <span>全部任务</span>
            <strong>{stats.total}</strong>
          </article>
          <article>
            <span>待完成</span>
            <strong>{stats.pending}</strong>
          </article>
          <article>
            <span>已完成</span>
            <strong>{stats.completed}</strong>
          </article>
          <article>
            <span>已逾期</span>
            <strong>{stats.overdue}</strong>
          </article>
        </section>

        <section className="content-grid">
          <form className="composer-card" onSubmit={handleSubmit}>
            <div className="section-heading">
              <h2>新建提醒</h2>
              <p>填写任务、说明和提醒时间。</p>
            </div>

            <label>
              <span>任务标题</span>
              <input
                type="text"
                placeholder="例如：提交日报"
                value={form.title}
                onChange={(event) => setForm((current) => ({ ...current, title: event.target.value }))}
              />
            </label>

            <label>
              <span>备注</span>
              <textarea
                rows="4"
                placeholder="补充任务说明、地点或准备事项"
                value={form.note}
                onChange={(event) => setForm((current) => ({ ...current, note: event.target.value }))}
              />
            </label>

            <div className="inline-fields">
              <label>
                <span>提醒时间</span>
                <input
                  type="datetime-local"
                  value={form.dueAt}
                  onChange={(event) => setForm((current) => ({ ...current, dueAt: event.target.value }))}
                />
              </label>

              <label>
                <span>优先级</span>
                <select
                  value={form.priority}
                  onChange={(event) => setForm((current) => ({ ...current, priority: event.target.value }))}
                >
                  <option value="high">高</option>
                  <option value="medium">中</option>
                  <option value="low">低</option>
                </select>
              </label>
            </div>

            <button className="primary-button" type="submit">
              添加代办
            </button>
          </form>

          <section className="list-card">
            <div className="list-toolbar">
              <div className="section-heading">
                <h2>任务列表</h2>
                <p>切换过滤条件，快速查看今日与逾期事项。</p>
              </div>
              <div className="filter-group">
                {[
                  ['all', '全部'],
                  ['today', '今天'],
                  ['pending', '待完成'],
                  ['completed', '已完成'],
                ].map(([value, label]) => (
                  <button
                    key={value}
                    type="button"
                    className={filter === value ? 'chip active' : 'chip'}
                    onClick={() => setFilter(value)}
                  >
                    {label}
                  </button>
                ))}
              </div>
            </div>

            <div className="task-list">
              {visibleTasks.length === 0 ? (
                <div className="empty-state">
                  <h3>当前没有匹配的任务</h3>
                  <p>新增一条提醒，或者切换筛选条件。</p>
                </div>
              ) : (
                visibleTasks.map((task) => {
                  const minutesLeft = getMinutesLeft(task.dueAt)
                  const overdue = !task.completed && minutesLeft < 0
                  return (
                    <article
                      className={task.completed ? 'task-item completed' : 'task-item'}
                      key={task.id}
                    >
                      <div className="task-main">
                        <div className="task-title-row">
                          <button
                            type="button"
                            className={task.completed ? 'toggle done' : 'toggle'}
                            onClick={() => toggleTask(task.id)}
                            aria-label={task.completed ? '标记为未完成' : '标记为已完成'}
                          />
                          <div>
                            <h3>{task.title}</h3>
                            <p>{task.note || '没有额外备注'}</p>
                          </div>
                        </div>
                        <div className="meta-row">
                          <span className={`priority-badge ${task.priority}`}>
                            {priorityLabel(task.priority)}
                          </span>
                          <span>{formatDateTime(task.dueAt)}</span>
                          <span className={overdue ? 'due-status overdue' : 'due-status'}>
                            {task.completed
                              ? '已完成'
                              : overdue
                                ? `已逾期 ${Math.abs(minutesLeft)} 分钟`
                                : `剩余 ${minutesLeft} 分钟`}
                          </span>
                        </div>
                      </div>
                      <button className="ghost-button" type="button" onClick={() => removeTask(task.id)}>
                        删除
                      </button>
                    </article>
                  )
                })
              )}
            </div>

            <div className="footer-actions">
              <button className="ghost-button" type="button" onClick={clearCompleted}>
                清除已完成任务
              </button>
            </div>
          </section>
        </section>
      </main>
    </div>
  )
}
