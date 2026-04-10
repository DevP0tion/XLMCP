# excel-mcp-server

PowerShell COM 자동화 기반 Excel MCP 서버. Claude Code 플러그인으로 동작.

열려 있는 Excel 인스턴스를 자동 감지하여 연결합니다. 워크북/시트 파라미터 생략 시 ActiveWorkbook/ActiveSheet를 사용합니다.

## 설치

```bash
bun install
```

## Claude Code 플러그인 등록

Claude Code 내에서:
```
/plugin install /path/to/excel-mcp-server
```

## 도구 목록 (18개)

### Workbook (5)
| 도구 | 설명 |
|---|---|
| `excel_list_open_workbooks` | 열린 워크북 목록 (이름, 경로, 시트 수) |
| `excel_get_active_workbook` | 활성 워크북 정보 |
| `excel_create_workbook` | 새 워크북 생성 |
| `excel_save_workbook` | 저장 / 다른 이름으로 저장 |
| `excel_close_workbook` | 워크북 닫기 |

### Sheet (5)
| 도구 | 설명 |
|---|---|
| `excel_list_sheets` | 시트 목록 |
| `excel_create_sheet` | 시트 추가 |
| `excel_delete_sheet` | 시트 삭제 |
| `excel_copy_sheet` | 시트 복사 |
| `excel_rename_sheet` | 시트 이름 변경 |

### Cell / Range (4)
| 도구 | 설명 |
|---|---|
| `excel_read_cell` | 단일 셀 읽기 (값, 수식, 표시 텍스트) |
| `excel_write_cell` | 단일 셀 쓰기 (값 또는 수식) |
| `excel_read_range` | 범위 읽기 → 2D 배열 (생략 시 UsedRange) |
| `excel_write_range` | 범위 쓰기 ← 2D 배열 |

### Format (4)
| 도구 | 설명 |
|---|---|
| `excel_format_range` | 서식 적용 (폰트, 색상, 정렬, 테두리, 표시 형식 등) |
| `excel_set_column_width` | 열 너비 (숫자 또는 auto) |
| `excel_set_row_height` | 행 높이 (숫자 또는 auto) |
| `excel_merge_cells` | 셀 병합 / 해제 |

## 구조

```
src/
├── index.ts              # MCP 서버 진입점 (stdio)
├── services/
│   ├── powershell.ts     # PowerShell COM 래퍼 + 워크북 자동 감지
│   └── utils.ts          # 유틸리티 (이스케이프, JSON 파싱 등)
├── tools/
│   ├── workbook.ts       # 워크북 관리 도구
│   ├── sheet.ts          # 시트 관리 도구
│   ├── cell.ts           # 셀/범위 읽기·쓰기 도구
│   └── format.ts         # 서식 도구
└── schemas/
    └── common.ts         # 공통 Zod 스키마
```
