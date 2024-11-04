import React from 'react'
import * as XLSX from 'xlsx'
import { saveAs } from 'file-saver'

interface ChannelData {
  title: string
  channelId: string
  subscriberCount: number
  dailyAverageViewCount: number
  viewRev: number
}

interface HeaderType {
  header: string
  key: string
}

export const ExcelTest: React.FC = () => {
  const sampleData: ChannelData[] = [
    {
      title: "테스트 채널1",
      channelId: "123",
      subscriberCount: 1000,
      dailyAverageViewCount: 500,
      viewRev: 100000
    },
    {
      title: "테스트 채널2",
      channelId: "456",
      subscriberCount: 2000,
      dailyAverageViewCount: 1000,
      viewRev: 200000
    }
  ]

  const downloadExcel = (): void => {
    try {
      const headers: HeaderType[] = [
        { header: 'Rank', key: 'Rank' },
        { header: 'Channel Name', key: 'title' },
        { header: 'Channel Link', key: 'channelLink' },
        { header: 'Subscribers', key: 'subscriberCount' },
        { header: 'Average Daily Views', key: 'dailyAverageViewCount' },
        { header: 'Estimated Revenue', key: 'viewRev' }
      ]

      const data = sampleData.map((channel, index) => {
        const rowData: { [key: string]: string | number } = {}
        headers.forEach(({ key, header }) => {
          if (key === 'Rank') {
            rowData[header] = index + 1
          } else if (key === 'channelLink') {
            rowData[header] = `https://test.com/channel/${channel.channelId}`
          } else {
            rowData[header] = (channel as any)[key] || ''
          }
        })
        return rowData
      })

      const ws = XLSX.utils.json_to_sheet(data, { 
        header: headers.map(h => h.header)
      })
      const wb = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(wb, ws, 'Channel Data')
      XLSX.writeFile(wb, 'test_download.xlsx')

    } catch (error) {
      console.error('Excel download error:', error)
    }
  }

  const downloadCSV = (): void => {
    try {
      const headers: string[] = ['Rank', 'Channel', 'Channel Link', 'Subscribers', 'Average Daily Views', 'Estimated Revenue']
      
      const rows = sampleData.map((channel, index) =>
        headers
          .map(header => {
            switch (header) {
              case 'Rank': return index + 1
              case 'Channel': return `"${channel.title}"`
              case 'Channel Link': return `"https://test.com/channel/${channel.channelId}"`
              case 'Subscribers': return channel.subscriberCount
              case 'Average Daily Views': return channel.dailyAverageViewCount
              case 'Estimated Revenue': return channel.viewRev
              default: return ''
            }
          })
          .join(',')
      )

      const csvContent = [headers.join(','), ...rows].join('\r\n')
      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' })
      saveAs(blob, 'test_download.csv')

    } catch (error) {
      console.error('CSV download error:', error)
    }
  }

  return (
    <div style={{ padding: '20px' }}>
      <h1>Excel/CSV Download Test</h1>
      <button onClick={downloadExcel} style={{ marginRight: '10px' }}>
        Excel 다운로드 테스트
      </button>
      <button onClick={downloadCSV}>
        CSV 다운로드 테스트
      </button>
    </div>
  )
}