import { useState, useEffect, useCallback, useMemo, useRef, type UIEvent } from 'react'
import { useNavigate } from 'react-router-dom'
import { Search, RefreshCw, X, User, Users, MessageSquare, Loader2, FolderOpen, Download, ChevronDown, MessageCircle, UserX } from 'lucide-react'
import { useChatStore } from '../stores/chatStore'
import { toContactTypeCardCounts, useContactTypeCountsStore } from '../stores/contactTypeCountsStore'
import './ContactsPage.scss'

interface ContactInfo {
    username: string
    displayName: string
    remark?: string
    nickname?: string
    avatarUrl?: string
    type: 'friend' | 'group' | 'official' | 'former_friend' | 'other'
}

interface ContactEnrichInfo {
    displayName?: string
    avatarUrl?: string
}

const AVATAR_ENRICH_BATCH_SIZE = 80
const SEARCH_DEBOUNCE_MS = 120
const VIRTUAL_ROW_HEIGHT = 76
const VIRTUAL_OVERSCAN = 10

function ContactsPage() {
    const [contacts, setContacts] = useState<ContactInfo[]>([])
    const [selectedUsernames, setSelectedUsernames] = useState<Set<string>>(new Set())
    const [isLoading, setIsLoading] = useState(true)
    const [searchKeyword, setSearchKeyword] = useState('')
    const [debouncedSearchKeyword, setDebouncedSearchKeyword] = useState('')
    const [contactTypes, setContactTypes] = useState({
        friends: true,
        groups: false,
        officials: false,
        deletedFriends: false
    })

    // 导出模式与查看详情
    const [exportMode, setExportMode] = useState(false)
    const [selectedContact, setSelectedContact] = useState<ContactInfo | null>(null)
    const navigate = useNavigate()
    const { setCurrentSession } = useChatStore()

    // 导出相关状态
    const [exportFormat, setExportFormat] = useState<'json' | 'csv' | 'vcf'>('json')
    const [exportAvatars, setExportAvatars] = useState(true)
    const [exportFolder, setExportFolder] = useState('')
    const [isExporting, setIsExporting] = useState(false)
    const [showFormatSelect, setShowFormatSelect] = useState(false)
    const formatDropdownRef = useRef<HTMLDivElement>(null)
    const listRef = useRef<HTMLDivElement>(null)
    const loadVersionRef = useRef(0)
    const [avatarEnrichProgress, setAvatarEnrichProgress] = useState({
        loaded: 0,
        total: 0,
        running: false
    })
    const [scrollTop, setScrollTop] = useState(0)
    const [listViewportHeight, setListViewportHeight] = useState(480)
    const sharedTabCounts = useContactTypeCountsStore(state => state.tabCounts)
    const syncContactTypeCounts = useContactTypeCountsStore(state => state.syncFromContacts)

    const applyEnrichedContacts = useCallback((enrichedMap: Record<string, ContactEnrichInfo>) => {
        if (!enrichedMap || Object.keys(enrichedMap).length === 0) return

        setContacts(prev => {
            let changed = false
            const next = prev.map(contact => {
                const enriched = enrichedMap[contact.username]
                if (!enriched) return contact
                const displayName = enriched.displayName || contact.displayName
                const avatarUrl = enriched.avatarUrl || contact.avatarUrl
                if (displayName === contact.displayName && avatarUrl === contact.avatarUrl) {
                    return contact
                }
                changed = true
                return {
                    ...contact,
                    displayName,
                    avatarUrl
                }
            })
            return changed ? next : prev
        })

        setSelectedContact(prev => {
            if (!prev) return prev
            const enriched = enrichedMap[prev.username]
            if (!enriched) return prev
            const displayName = enriched.displayName || prev.displayName
            const avatarUrl = enriched.avatarUrl || prev.avatarUrl
            if (displayName === prev.displayName && avatarUrl === prev.avatarUrl) {
                return prev
            }
            return {
                ...prev,
                displayName,
                avatarUrl
            }
        })
    }, [])

    const enrichContactsInBackground = useCallback(async (sourceContacts: ContactInfo[], loadVersion: number) => {
        const usernames = sourceContacts.map(contact => contact.username).filter(Boolean)
        const total = usernames.length
        setAvatarEnrichProgress({
            loaded: 0,
            total,
            running: total > 0
        })
        if (total === 0) return

        for (let i = 0; i < total; i += AVATAR_ENRICH_BATCH_SIZE) {
            if (loadVersionRef.current !== loadVersion) return
            const batch = usernames.slice(i, i + AVATAR_ENRICH_BATCH_SIZE)
            if (batch.length === 0) continue

            try {
                const avatarResult = await window.electronAPI.chat.enrichSessionsContactInfo(batch)
                if (loadVersionRef.current !== loadVersion) return
                if (avatarResult.success && avatarResult.contacts) {
                    applyEnrichedContacts(avatarResult.contacts)
                }
            } catch (e) {
                console.error('分批补全头像失败:', e)
            }

            const loaded = Math.min(i + batch.length, total)
            setAvatarEnrichProgress({
                loaded,
                total,
                running: loaded < total
            })

            await new Promise(resolve => setTimeout(resolve, 0))
        }
    }, [applyEnrichedContacts])

    // 加载通讯录
    const loadContacts = useCallback(async () => {
        const loadVersion = loadVersionRef.current + 1
        loadVersionRef.current = loadVersion
        setIsLoading(true)
        setAvatarEnrichProgress({
            loaded: 0,
            total: 0,
            running: false
        })
        try {
            const contactsResult = await window.electronAPI.chat.getContacts()

            if (loadVersionRef.current !== loadVersion) return
            if (contactsResult.success && contactsResult.contacts) {
                setContacts(contactsResult.contacts)
                syncContactTypeCounts(contactsResult.contacts)
                setSelectedUsernames(new Set())
                setSelectedContact(prev => {
                    if (!prev) return prev
                    return contactsResult.contacts!.find(contact => contact.username === prev.username) || null
                })
                setIsLoading(false)
                void enrichContactsInBackground(contactsResult.contacts, loadVersion)
                return
            }
        } catch (e) {
            console.error('加载通讯录失败:', e)
        } finally {
            if (loadVersionRef.current === loadVersion) {
                setIsLoading(false)
            }
        }
    }, [enrichContactsInBackground, syncContactTypeCounts])

    useEffect(() => {
        loadContacts()
    }, [loadContacts])

    useEffect(() => {
        return () => {
            loadVersionRef.current += 1
        }
    }, [])

    useEffect(() => {
        const timer = window.setTimeout(() => {
            setDebouncedSearchKeyword(searchKeyword.trim().toLowerCase())
        }, SEARCH_DEBOUNCE_MS)
        return () => window.clearTimeout(timer)
    }, [searchKeyword])

    const filteredContacts = useMemo(() => {
        let filtered = contacts.filter(contact => {
            if (contact.type === 'friend' && !contactTypes.friends) return false
            if (contact.type === 'group' && !contactTypes.groups) return false
            if (contact.type === 'official' && !contactTypes.officials) return false
            if (contact.type === 'former_friend' && !contactTypes.deletedFriends) return false
            return true
        })

        if (debouncedSearchKeyword) {
            filtered = filtered.filter(contact =>
                contact.displayName?.toLowerCase().includes(debouncedSearchKeyword) ||
                contact.remark?.toLowerCase().includes(debouncedSearchKeyword) ||
                contact.username.toLowerCase().includes(debouncedSearchKeyword)
            )
        }

        return filtered
    }, [contacts, contactTypes, debouncedSearchKeyword])

    const contactTypeCounts = useMemo(() => toContactTypeCardCounts(sharedTabCounts), [sharedTabCounts])

    useEffect(() => {
        if (!listRef.current) return
        listRef.current.scrollTop = 0
        setScrollTop(0)
    }, [debouncedSearchKeyword, contactTypes])

    useEffect(() => {
        const node = listRef.current
        if (!node) return

        const updateViewportHeight = () => {
            setListViewportHeight(Math.max(node.clientHeight, VIRTUAL_ROW_HEIGHT))
        }
        updateViewportHeight()

        const observer = new ResizeObserver(() => updateViewportHeight())
        observer.observe(node)
        return () => observer.disconnect()
    }, [filteredContacts.length, isLoading])

    useEffect(() => {
        const maxScroll = Math.max(0, filteredContacts.length * VIRTUAL_ROW_HEIGHT - listViewportHeight)
        if (scrollTop <= maxScroll) return
        setScrollTop(maxScroll)
        if (listRef.current) {
            listRef.current.scrollTop = maxScroll
        }
    }, [filteredContacts.length, listViewportHeight, scrollTop])

    // 搜索和类型过滤
    useEffect(() => {
        const handleClickOutside = (event: MouseEvent) => {
            const target = event.target as Node
            if (showFormatSelect && formatDropdownRef.current && !formatDropdownRef.current.contains(target)) {
                setShowFormatSelect(false)
            }
        }
        document.addEventListener('mousedown', handleClickOutside)
        return () => document.removeEventListener('mousedown', handleClickOutside)
    }, [showFormatSelect])

    const selectedInFilteredCount = useMemo(() => {
        return filteredContacts.reduce((count, contact) => {
            return selectedUsernames.has(contact.username) ? count + 1 : count
        }, 0)
    }, [filteredContacts, selectedUsernames])
    const allFilteredSelected = filteredContacts.length > 0 && selectedInFilteredCount === filteredContacts.length

    const { startIndex, endIndex } = useMemo(() => {
        if (filteredContacts.length === 0) {
            return { startIndex: 0, endIndex: 0 }
        }
        const baseStart = Math.floor(scrollTop / VIRTUAL_ROW_HEIGHT)
        const visibleCount = Math.ceil(listViewportHeight / VIRTUAL_ROW_HEIGHT)
        const nextStart = Math.max(0, baseStart - VIRTUAL_OVERSCAN)
        const nextEnd = Math.min(filteredContacts.length, nextStart + visibleCount + VIRTUAL_OVERSCAN * 2)
        return {
            startIndex: nextStart,
            endIndex: nextEnd
        }
    }, [filteredContacts.length, listViewportHeight, scrollTop])

    const visibleContacts = useMemo(() => {
        return filteredContacts.slice(startIndex, endIndex)
    }, [filteredContacts, startIndex, endIndex])

    const onContactsListScroll = useCallback((event: UIEvent<HTMLDivElement>) => {
        setScrollTop(event.currentTarget.scrollTop)
    }, [])

    const toggleContactSelected = (username: string, checked: boolean) => {
        setSelectedUsernames(prev => {
            const next = new Set(prev)
            if (checked) {
                next.add(username)
            } else {
                next.delete(username)
            }
            return next
        })
    }

    const toggleAllFilteredSelected = (checked: boolean) => {
        setSelectedUsernames(prev => {
            const next = new Set(prev)
            filteredContacts.forEach(contact => {
                if (checked) {
                    next.add(contact.username)
                } else {
                    next.delete(contact.username)
                }
            })
            return next
        })
    }

    const getAvatarLetter = (name: string) => {
        if (!name) return '?'
        return [...name][0] || '?'
    }

    const getContactTypeIcon = (type: string) => {
        switch (type) {
            case 'friend': return <User size={14} />
            case 'group': return <Users size={14} />
            case 'official': return <MessageSquare size={14} />
            case 'former_friend': return <UserX size={14} />
            default: return <User size={14} />
        }
    }

    const getContactTypeName = (type: string) => {
        switch (type) {
            case 'friend': return '好友'
            case 'group': return '群聊'
            case 'official': return '公众号'
            case 'former_friend': return '曾经的好友'
            default: return '其他'
        }
    }

    // 选择导出文件夹
    const selectExportFolder = async () => {
        try {
            const result = await window.electronAPI.dialog.openDirectory({
                title: '选择导出位置'
            })
            if (result && !result.canceled && result.filePaths && result.filePaths.length > 0) {
                setExportFolder(result.filePaths[0])
            }
        } catch (e) {
            console.error('选择文件夹失败:', e)
        }
    }

    // 开始导出
    const startExport = async () => {
        if (!exportFolder) {
            alert('请先选择导出位置')
            return
        }
        if (selectedUsernames.size === 0) {
            alert('请至少选择一个联系人')
            return
        }

        setIsExporting(true)
        try {
            const exportOptions = {
                format: exportFormat,
                exportAvatars,
                contactTypes: {
                    friends: contactTypes.friends,
                    groups: contactTypes.groups,
                    officials: contactTypes.officials
                },
                selectedUsernames: Array.from(selectedUsernames)
            }

            const result = await window.electronAPI.export.exportContacts(exportFolder, exportOptions)

            if (result.success) {
                alert(`导出成功！共导出 ${result.successCount} 个联系人`)
            } else {
                alert(`导出失败：${result.error}`)
            }
        } catch (e) {
            console.error('导出失败:', e)
            alert(`导出失败：${String(e)}`)
        } finally {
            setIsExporting(false)
        }
    }

    const exportFormatOptions = [
        { value: 'json', label: 'JSON', desc: '详细格式，包含完整联系人信息' },
        { value: 'csv', label: 'CSV (Excel)', desc: '电子表格格式，适合Excel查看' },
        { value: 'vcf', label: 'VCF (vCard)', desc: '标准名片格式，支持导入手机' }
    ]

    const getOptionLabel = (value: string) => {
        return exportFormatOptions.find(opt => opt.value === value)?.label || value
    }

    return (
        <div className="contacts-page">
            {/* 左侧：联系人列表 */}
            <div className="contacts-panel">
                <div className="panel-header">
                    <h2>通讯录</h2>
                    <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                        <button
                            className={`icon-btn export-mode-btn ${exportMode ? 'active' : ''}`}
                            onClick={() => { setExportMode(!exportMode); setSelectedContact(null) }}
                            title={exportMode ? '退出导出模式' : '进入导出模式'}
                        >
                            <Download size={18} />
                        </button>
                        <button className="icon-btn" onClick={loadContacts} disabled={isLoading}>
                            <RefreshCw size={18} className={isLoading ? 'spin' : ''} />
                        </button>
                    </div>
                </div>

                <div className="search-bar">
                    <Search size={16} />
                    <input
                        type="text"
                        placeholder="搜索联系人..."
                        value={searchKeyword}
                        onChange={e => setSearchKeyword(e.target.value)}
                    />
                    {searchKeyword && (
                        <button className="clear-btn" onClick={() => setSearchKeyword('')}>
                            <X size={14} />
                        </button>
                    )}
                </div>

                <div className="type-filters">
                    <label className={`filter-chip ${contactTypes.friends ? 'active' : ''}`}>
                        <input type="checkbox" checked={contactTypes.friends} onChange={e => setContactTypes({ ...contactTypes, friends: e.target.checked })} />
                        <User size={16} />
                        <span className="chip-label">好友</span>
                        <span className="chip-count">{contactTypeCounts.friends}</span>
                    </label>
                    <label className={`filter-chip ${contactTypes.groups ? 'active' : ''}`}>
                        <input type="checkbox" checked={contactTypes.groups} onChange={e => setContactTypes({ ...contactTypes, groups: e.target.checked })} />
                        <Users size={16} />
                        <span className="chip-label">群聊</span>
                        <span className="chip-count">{contactTypeCounts.groups}</span>
                    </label>
                    <label className={`filter-chip ${contactTypes.officials ? 'active' : ''}`}>
                        <input type="checkbox" checked={contactTypes.officials} onChange={e => setContactTypes({ ...contactTypes, officials: e.target.checked })} />
                        <MessageSquare size={16} />
                        <span className="chip-label">公众号</span>
                        <span className="chip-count">{contactTypeCounts.officials}</span>
                    </label>
                    <label className={`filter-chip ${contactTypes.deletedFriends ? 'active' : ''}`}>
                        <input type="checkbox" checked={contactTypes.deletedFriends} onChange={e => setContactTypes({ ...contactTypes, deletedFriends: e.target.checked })} />
                        <UserX size={16} />
                        <span className="chip-label">曾经的好友</span>
                        <span className="chip-count">{contactTypeCounts.deletedFriends}</span>
                    </label>
                </div>

                <div className="contacts-count">
                    共 {filteredContacts.length} / {contacts.length} 个联系人
                    {avatarEnrichProgress.running && (
                        <span className="avatar-enrich-progress">
                            头像补全中 {avatarEnrichProgress.loaded}/{avatarEnrichProgress.total}
                        </span>
                    )}
                </div>

                {exportMode && (
                    <div className="selection-toolbar">
                        <label className="checkbox-item">
                            <input
                                type="checkbox"
                                checked={allFilteredSelected}
                                onChange={e => toggleAllFilteredSelected(e.target.checked)}
                                disabled={filteredContacts.length === 0}
                            />
                            <span>全选当前筛选结果</span>
                        </label>
                        <span className="selection-count">已选 {selectedUsernames.size}（当前筛选 {selectedInFilteredCount} / {filteredContacts.length}）</span>
                    </div>
                )}

                {isLoading && contacts.length === 0 ? (
                    <div className="loading-state">
                        <Loader2 size={32} className="spin" />
                        <span>联系人加载中...</span>
                    </div>
                ) : filteredContacts.length === 0 ? (
                    <div className="empty-state">
                        <span>暂无联系人</span>
                    </div>
                ) : (
                    <div className="contacts-list" ref={listRef} onScroll={onContactsListScroll}>
                        <div
                            className="contacts-list-virtual"
                            style={{ height: filteredContacts.length * VIRTUAL_ROW_HEIGHT }}
                        >
                            {visibleContacts.map((contact, idx) => {
                            const absoluteIndex = startIndex + idx
                            const top = absoluteIndex * VIRTUAL_ROW_HEIGHT
                            const isChecked = selectedUsernames.has(contact.username)
                            const isActive = !exportMode && selectedContact?.username === contact.username
                            return (
                                <div
                                    key={contact.username}
                                    className="contact-row"
                                    style={{ transform: `translateY(${top}px)` }}
                                >
                                    <div
                                        className={`contact-item ${exportMode && isChecked ? 'selected' : ''} ${isActive ? 'active' : ''}`}
                                        onClick={() => {
                                            if (exportMode) {
                                                toggleContactSelected(contact.username, !isChecked)
                                            } else {
                                                setSelectedContact(isActive ? null : contact)
                                            }
                                        }}
                                    >
                                        {exportMode && (
                                            <label className="contact-select" onClick={e => e.stopPropagation()}>
                                                <input
                                                    type="checkbox"
                                                    checked={isChecked}
                                                    onChange={e => toggleContactSelected(contact.username, e.target.checked)}
                                                />
                                            </label>
                                        )}
                                        <div className="contact-avatar">
                                            {contact.avatarUrl ? (
                                                <img src={contact.avatarUrl} alt="" loading="lazy" />
                                            ) : (
                                                <span>{getAvatarLetter(contact.displayName)}</span>
                                            )}
                                        </div>
                                        <div className="contact-info">
                                            <div className="contact-name">{contact.displayName}</div>
                                            {contact.remark && contact.remark !== contact.displayName && (
                                                <div className="contact-remark">备注: {contact.remark}</div>
                                            )}
                                        </div>
                                        <div className={`contact-type ${contact.type}`}>
                                            {getContactTypeIcon(contact.type)}
                                            <span>{getContactTypeName(contact.type)}</span>
                                        </div>
                                    </div>
                                </div>
                            )
                            })}
                        </div>
                    </div>
                )}
            </div>

            {/* 右侧面板 */}
            {exportMode ? (
                <div className="settings-panel">
                    <div className="panel-header">
                        <h2>导出设置</h2>
                    </div>

                    <div className="settings-content">
                        <div className="setting-section">
                            <h3>导出格式</h3>
                            <div className="format-select" ref={formatDropdownRef}>
                                <button
                                    type="button"
                                    className={`select-trigger ${showFormatSelect ? 'open' : ''}`}
                                    onClick={() => setShowFormatSelect(!showFormatSelect)}
                                >
                                    <span className="select-value">{getOptionLabel(exportFormat)}</span>
                                    <ChevronDown size={16} />
                                </button>
                                {showFormatSelect && (
                                    <div className="select-dropdown">
                                        {exportFormatOptions.map(option => (
                                            <button
                                                key={option.value}
                                                type="button"
                                                className={`select-option ${exportFormat === option.value ? 'active' : ''}`}
                                                onClick={() => {
                                                    setExportFormat(option.value as 'json' | 'csv' | 'vcf')
                                                    setShowFormatSelect(false)
                                                }}
                                            >
                                                <span className="option-label">{option.label}</span>
                                                <span className="option-desc">{option.desc}</span>
                                            </button>
                                        ))}
                                    </div>
                                )}
                            </div>
                        </div>

                        <div className="setting-section">
                            <h3>导出选项</h3>
                            <label className="checkbox-item">
                                <input type="checkbox" checked={exportAvatars} onChange={e => setExportAvatars(e.target.checked)} />
                                <span>导出头像</span>
                            </label>
                        </div>

                        <div className="setting-section">
                            <h3>导出位置</h3>
                            <div className="export-path-display">
                                <FolderOpen size={16} />
                                <span>{exportFolder || '未设置'}</span>
                            </div>
                            <button className="select-folder-btn" onClick={selectExportFolder}>
                                <FolderOpen size={16} />
                                <span>选择导出目录</span>
                            </button>
                        </div>
                    </div>

                    <div className="export-action">
                        <button
                            className="export-btn"
                            onClick={startExport}
                            disabled={!exportFolder || isExporting || selectedUsernames.size === 0}
                        >
                            {isExporting ? (
                                <><Loader2 size={18} className="spin" /><span>导出中...</span></>
                            ) : (
                                <><Download size={18} /><span>开始导出</span></>
                            )}
                        </button>
                    </div>
                </div>
            ) : selectedContact ? (
                <div className="settings-panel">
                    <div className="panel-header">
                        <h2>联系人详情</h2>
                    </div>
                    <div className="settings-content">
                        <div className="detail-profile">
                            <div className="detail-avatar">
                                {selectedContact.avatarUrl ? (
                                    <img src={selectedContact.avatarUrl} alt="" />
                                ) : (
                                    <span>{getAvatarLetter(selectedContact.displayName)}</span>
                                )}
                            </div>
                            <div className="detail-name">{selectedContact.displayName}</div>
                            <div className={`contact-type ${selectedContact.type}`}>
                                {getContactTypeIcon(selectedContact.type)}
                                <span>{getContactTypeName(selectedContact.type)}</span>
                            </div>
                        </div>

                        <div className="detail-info-list">
                            <div className="detail-row"><span className="detail-label">用户名</span><span className="detail-value">{selectedContact.username}</span></div>
                            <div className="detail-row"><span className="detail-label">昵称</span><span className="detail-value">{selectedContact.nickname || selectedContact.displayName}</span></div>
                            {selectedContact.remark && <div className="detail-row"><span className="detail-label">备注</span><span className="detail-value">{selectedContact.remark}</span></div>}
                            <div className="detail-row"><span className="detail-label">类型</span><span className="detail-value">{getContactTypeName(selectedContact.type)}</span></div>
                        </div>

                        <button
                            className="goto-chat-btn"
                            onClick={() => {
                                setCurrentSession(selectedContact.username)
                                navigate('/chat')
                            }}
                        >
                            <MessageCircle size={18} />
                            <span>查看聊天记录</span>
                        </button>
                    </div>
                </div>
            ) : (
                <div className="settings-panel">
                    <div className="empty-detail">
                        <User size={48} />
                        <span>点击左侧联系人查看详情</span>
                    </div>
                </div>
            )}
        </div>
    )
}

export default ContactsPage
