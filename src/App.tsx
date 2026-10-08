/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { 
  CheckCircle2, 
  Clock, 
  Calendar, 
  Tag, 
  Trash2, 
  RotateCcw, 
  PlusCircle, 
  Check, 
  FileText,
  AlertCircle,
  FileSpreadsheet
} from 'lucide-react';

/* ========================================================
   [타입 정의 (TypeScript Interface)]
   - 과제 데이터 모델 인터페이스 정의
   ======================================================== */
interface Task {
  id: string;              // 과제 고유 식별자 ID
  name: string;            // 과제 이름
  duration: number;        // 소요시간 (1~24시간)
  category: string;        // 분류: 개인과제, 팀플, 발표
  priority: '하' | '중' | '상'; // 중요도
  options: string[];       // 추가 옵션: 코딩, 디자인/아트, 교양, 언어
  dueDate: string;         // 날짜
  memo: string;            // 메모사항
  isCompleted: boolean;    // 완료 여부 상태값
  createdAt: number;       // 등록 시간 타임스탬프
}

export default function App() {
  /* ========================================================
     [상태 관리 (State)]
     - 폼 입력 데이터 상태 관리
     ======================================================== */
  const [taskName, setTaskName] = useState<string>('');
  const [durationInput, setDurationInput] = useState<string>(''); // 소요시간 입력값 (실시간 반영용)
  const [category, setCategory] = useState<string>('개인과제');
  const [priority, setPriority] = useState<'하' | '중' | '상'>('중'); // 기본 선택: '중'
  const [selectedOptions, setSelectedOptions] = useState<string[]>([]);
  // 기본 날짜는 오늘 날짜(YYYY-MM-DD)로 기본 설정
  const [dueDate, setDueDate] = useState<string>(() => {
    const today = new Date();
    return today.toISOString().split('T')[0];
  });
  const [memo, setMemo] = useState<string>('');

  /* ========================================================
     [게시판 목록 상태 관리 (Task List State)]
     - 등록된 과제 카드들의 리스트 관리
     - 초기 데모 과제 2개를 기본 제공하여 첫 실행 시에도 완성도를 확인 가능하도록 구성
     ======================================================== */
  const [taskList, setTaskList] = useState<Task[]>([
    {
      id: 'demo-task-1',
      name: '웹 프로그래밍 과제 제출하기',
      duration: 3,
      category: '개인과제',
      priority: '상',
      options: ['코딩'],
      dueDate: new Date().toISOString().split('T')[0],
      memo: 'GitHub 저장소 링크 및 리포트 작성 후 LMS 업로드',
      isCompleted: false,
      createdAt: Date.now() - 100000,
    },
    {
      id: 'demo-task-2',
      name: '소프트웨어 공학 팀 회의 준비',
      duration: 2,
      category: '팀플',
      priority: '중',
      options: ['교양'],
      dueDate: new Date().toISOString().split('T')[0],
      memo: '요구사항 정의서 검토 및 발표 슬라이드 초안 구성',
      isCompleted: true,
      createdAt: Date.now() - 200000,
    }
  ]);

  /* ========================================================
     [실시간 계산 및 통계]
     - 실시간 예상 소요시간 계산: 숫자가 유효하면 해당 시간 표시, 아니면 0시간
     - 미완료 과제 개수 실시간 카운트
     ======================================================== */
  const parsedDuration = parseInt(durationInput, 10);
  const displayDuration = !isNaN(parsedDuration) && parsedDuration > 0 ? parsedDuration : 0;

  // 완료되지 않은 과제 개수 계산
  const uncompletedCount = taskList.filter((task) => !task.isCompleted).length;

  const handleDownloadExcel = () => {
    const rows = taskList.map((task) => ({
      과제명: task.name,
      분류: task.category,
      중요도: task.priority,
      소요시간: task.duration,
      마감일: task.dueDate,
      옵션: task.options.join(', '),
      메모: task.memo,
      상태: task.isCompleted ? '완료' : '진행 중',
    }));
    const worksheet = XLSX.utils.json_to_sheet(rows);
    worksheet['!cols'] = [
      { wch: 32 }, { wch: 14 }, { wch: 10 }, { wch: 12 },
      { wch: 14 }, { wch: 24 }, { wch: 48 }, { wch: 12 },
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '과제 목록');
    XLSX.writeFile(workbook, `과제목록_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  /* ========================================================
     [핸들러 1: 추가 옵션 체크박스 변경]
     - 체크된 항목을 배열에 추가/제거
     ======================================================== */
  const handleOptionChange = (optionValue: string) => {
    setSelectedOptions((prev) =>
      prev.includes(optionValue)
        ? prev.filter((item) => item !== optionValue)
        : [...prev, optionValue]
    );
  };

  /* ========================================================
     [핸들러 2: 과제 등록하기 버튼 클릭 이벤트]
     - 유효성 검사 수행 (이름 필수, 소요시간 1~24 범위)
     - 정상이면 신규 과제 카드 생성 후 목록에 추가 및 폼 초기화
     ======================================================== */
  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();

    // 1. 과제 이름 필수 검사 (공백 제외)
    if (!taskName.trim()) {
      alert('과제 이름을 입력해주세요');
      return;
    }

    // 2. 소요시간 유효성 검사 (1~24시간 사이)
    const durationNum = Number(durationInput);
    if (!durationInput || isNaN(durationNum) || durationNum < 1 || durationNum > 24) {
      alert('소요시간을 1~24시간 사이로 입력해주세요');
      return;
    }

    // 3. 신규 과제 객체 생성
    const newTask: Task = {
      id: `task-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`,
      name: taskName.trim(),
      duration: durationNum,
      category,
      priority,
      options: [...selectedOptions],
      dueDate: dueDate || new Date().toISOString().split('T')[0],
      memo: memo.trim(),
      isCompleted: false, // 새로 등록된 과제는 기본 미완료 상태
      createdAt: Date.now(),
    };

    // 4. 게시판 영역에 동적으로 새 과제 카드 추가
    setTaskList((prev) => [newTask, ...prev]);

    // 5. 폼 입력값 초기화
    handleReset();
  };

  /* ========================================================
     [핸들러 3: 다시 작성 버튼 (폼 초기화)]
     - 현재 폼에 입력 중인 모든 정보와 예상 소요시간 표시 초기화
     ======================================================== */
  const handleReset = () => {
    setTaskName('');
    setDurationInput('');
    setCategory('개인과제');
    setPriority('중'); // 기본 선택값 '중'으로 복원
    setSelectedOptions([]);
    const today = new Date().toISOString().split('T')[0];
    setDueDate(today);
    setMemo('');
  };

  /* ========================================================
     [핸들러 4: 과제 완료 토글 버튼]
     - 클릭 시 완료 상태 토글 (취소선 및 어두운 배경 적용, 미완료 수 실시간 반영)
     ======================================================== */
  const handleToggleComplete = (taskId: string) => {
    setTaskList((prevList) =>
      prevList.map((task) =>
        task.id === taskId ? { ...task, isCompleted: !task.isCompleted } : task
      )
    );
  };

  /* ========================================================
     [핸들러 5: 과제 카드 삭제]
     - 사용자가 개별 과제를 제거할 수 있는 기능
     ======================================================== */
  const handleDeleteTask = (taskId: string) => {
    if (window.confirm('이 과제를 삭제하시겠습니까?')) {
      setTaskList((prevList) => prevList.filter((task) => task.id !== taskId));
    }
  };

  return (
    <div className="min-h-screen py-8 px-4 flex justify-center items-start">
      {/* 
        [전체 레이아웃 컨테이너]
        - 최대 너비 520px, 가운데 정렬 (max-w-[520px] mx-auto)
        - IT 느낌의 깔끔한 색상 및 둥근 모서리, 부드러운 그림자
      */}
      <main className="w-full max-w-[520px] mx-auto flex flex-col gap-6">

        {/* ========================================================
            [페이지 상단 헤더 영역]
            - 로고: 컴퓨터 이모지 크게
            - 이름: "과제 게시판"
            - 부제: "당신의 하루에 과제를 더하다"
            ======================================================== */}
        <header className="bg-white rounded-2xl p-6 text-center shadow-md border border-blue-100 flex flex-col items-center">
          {/* 컴퓨터 이모지 로고 */}
          <div className="w-16 h-16 rounded-2xl bg-blue-50 border border-blue-200 flex items-center justify-center text-4xl shadow-inner mb-3 transform hover:scale-105 transition-transform duration-200" aria-label="컴퓨터 로고">
            💻
          </div>
          {/* 게시판 이름 */}
          <h1 className="text-2xl font-black text-slate-900 tracking-tight">
            과제 게시판
          </h1>
          {/* 부제 */}
          <p className="text-sm font-medium text-blue-600 mt-1">
            당신의 하루에 과제를 더하다
          </p>
        </header>

        {/* ========================================================
            [과제 등록 폼 섹션]
            - IT 느낌의 파랑/하늘/화이트 톤 둥근 카드
            - 모든 input에 label id와 htmlFor 정확히 연결
            ======================================================== */}
        <section 
          aria-labelledby="form-heading" 
          className="bg-white rounded-2xl p-6 shadow-md border border-slate-200"
        >
          <div className="flex items-center gap-2 mb-5 pb-3 border-b border-slate-100">
            <PlusCircle className="w-5 h-5 text-blue-600" />
            <h2 id="form-heading" className="text-lg font-bold text-slate-800">
              새로운 과제 등록
            </h2>
          </div>

          <form onSubmit={handleSubmit} className="flex flex-col gap-4">

            {/* [1. 과제 이름] (필수, text) */}
            <div className="flex flex-col gap-1.5">
              <label 
                htmlFor="taskName" 
                className="text-sm font-bold text-slate-700 flex items-center gap-1"
              >
                1. 과제 이름 <span className="text-rose-500 font-bold">* (필수)</span>
              </label>
              <input
                id="taskName"
                name="taskName"
                type="text"
                placeholder="예: 데이터베이스 정규화 과제"
                value={taskName}
                onChange={(e) => setTaskName(e.target.value)}
                className="custom-input text-sm text-slate-900 placeholder:text-slate-400"
              />
            </div>

            {/* [2. 소요시간] (min 1h, max 24h, 실시간 반영) */}
            <div className="flex flex-col gap-1.5">
              <label 
                htmlFor="taskDuration" 
                className="text-sm font-bold text-slate-700 flex items-center gap-1"
              >
                <Clock className="w-4 h-4 text-blue-600" />
                2. 소요시간 <span className="text-xs text-slate-500 font-normal">(1 ~ 24시간)</span>
              </label>
              <input
                id="taskDuration"
                name="taskDuration"
                type="number"
                min="1"
                max="24"
                placeholder="예: 3 (1~24 사이의 숫자 입력)"
                value={durationInput}
                onChange={(e) => setDurationInput(e.target.value)}
                className="custom-input text-sm text-slate-900 placeholder:text-slate-400"
              />
            </div>

            {/* [3. 분류 선택] (드롭다운) */}
            <div className="flex flex-col gap-1.5">
              <label 
                htmlFor="taskCategory" 
                className="text-sm font-bold text-slate-700 flex items-center gap-1"
              >
                <Tag className="w-4 h-4 text-blue-600" />
                3. 분류 선택
              </label>
              <select
                id="taskCategory"
                name="taskCategory"
                value={category}
                onChange={(e) => setCategory(e.target.value)}
                className="custom-input text-sm text-slate-900 cursor-pointer bg-white"
              >
                <option value="개인과제">개인과제</option>
                <option value="팀플">팀플</option>
                <option value="발표">발표</option>
              </select>
            </div>

            {/* [4. 중요도] (라디오 버튼, 가로 배치, 기본 선택: '중') */}
            <div className="flex flex-col gap-1.5">
              <span className="text-sm font-bold text-slate-700">
                4. 중요도 <span className="text-xs text-slate-500 font-normal">(기본: 중)</span>
              </span>
              <div className="flex flex-row items-center gap-4 py-1">
                {/* 중요도: 하 */}
                <div className="flex items-center gap-1.5">
                  <input
                    type="radio"
                    id="priorityLow"
                    name="priority"
                    value="하"
                    checked={priority === '하'}
                    onChange={() => setPriority('하')}
                    className="w-4 h-4 text-blue-600 cursor-pointer accent-blue-600"
                  />
                  <label htmlFor="priorityLow" className="text-sm font-medium text-slate-700 cursor-pointer select-none">
                    하
                  </label>
                </div>

                {/* 중요도: 중 (기본 선택) */}
                <div className="flex items-center gap-1.5">
                  <input
                    type="radio"
                    id="priorityMed"
                    name="priority"
                    value="중"
                    checked={priority === '중'}
                    onChange={() => setPriority('중')}
                    className="w-4 h-4 text-blue-600 cursor-pointer accent-blue-600"
                  />
                  <label htmlFor="priorityMed" className="text-sm font-medium text-slate-700 cursor-pointer select-none">
                    중
                  </label>
                </div>

                {/* 중요도: 상 */}
                <div className="flex items-center gap-1.5">
                  <input
                    type="radio"
                    id="priorityHigh"
                    name="priority"
                    value="상"
                    checked={priority === '상'}
                    onChange={() => setPriority('상')}
                    className="w-4 h-4 text-blue-600 cursor-pointer accent-blue-600"
                  />
                  <label htmlFor="priorityHigh" className="text-sm font-medium text-slate-700 cursor-pointer select-none">
                    상
                  </label>
                </div>
              </div>
            </div>

            {/* [5. 추가 옵션] (체크박스, 가로 배치) */}
            <div className="flex flex-col gap-1.5">
              <span className="text-sm font-bold text-slate-700">
                5. 추가 옵션 <span className="text-xs text-slate-500 font-normal">(중복 선택 가능)</span>
              </span>
              <div className="flex flex-row flex-wrap items-center gap-4 py-1">
                {/* 코딩 */}
                <div className="flex items-center gap-1.5">
                  <input
                    type="checkbox"
                    id="optCoding"
                    name="options"
                    value="코딩"
                    checked={selectedOptions.includes('코딩')}
                    onChange={() => handleOptionChange('코딩')}
                    className="w-4 h-4 rounded text-blue-600 cursor-pointer accent-blue-600"
                  />
                  <label htmlFor="optCoding" className="text-sm font-medium text-slate-700 cursor-pointer select-none">
                    코딩
                  </label>
                </div>

                {/* 디자인/아트 */}
                <div className="flex items-center gap-1.5">
                  <input
                    type="checkbox"
                    id="optDesign"
                    name="options"
                    value="디자인/아트"
                    checked={selectedOptions.includes('디자인/아트')}
                    onChange={() => handleOptionChange('디자인/아트')}
                    className="w-4 h-4 rounded text-blue-600 cursor-pointer accent-blue-600"
                  />
                  <label htmlFor="optDesign" className="text-sm font-medium text-slate-700 cursor-pointer select-none">
                    디자인/아트
                  </label>
                </div>

                {/* 교양 */}
                <div className="flex items-center gap-1.5">
                  <input
                    type="checkbox"
                    id="optLiberal"
                    name="options"
                    value="교양"
                    checked={selectedOptions.includes('교양')}
                    onChange={() => handleOptionChange('교양')}
                    className="w-4 h-4 rounded text-blue-600 cursor-pointer accent-blue-600"
                  />
                  <label htmlFor="optLiberal" className="text-sm font-medium text-slate-700 cursor-pointer select-none">
                    교양
                  </label>
                </div>

                {/* 언어 */}
                <div className="flex items-center gap-1.5">
                  <input
                    type="checkbox"
                    id="optLang"
                    name="options"
                    value="언어"
                    checked={selectedOptions.includes('언어')}
                    onChange={() => handleOptionChange('언어')}
                    className="w-4 h-4 rounded text-blue-600 cursor-pointer accent-blue-600"
                  />
                  <label htmlFor="optLang" className="text-sm font-medium text-slate-700 cursor-pointer select-none">
                    언어
                  </label>
                </div>
              </div>
            </div>

            {/* [6. 날짜] */}
            <div className="flex flex-col gap-1.5">
              <label 
                htmlFor="taskDueDate" 
                className="text-sm font-bold text-slate-700 flex items-center gap-1"
              >
                <Calendar className="w-4 h-4 text-blue-600" />
                6. 날짜
              </label>
              <input
                id="taskDueDate"
                name="taskDueDate"
                type="date"
                value={dueDate}
                onChange={(e) => setDueDate(e.target.value)}
                className="custom-input text-sm text-slate-900 cursor-pointer"
              />
            </div>

            {/* [7. 메모사항] (textarea) */}
            <div className="flex flex-col gap-1.5">
              <label 
                htmlFor="taskMemo" 
                className="text-sm font-bold text-slate-700 flex items-center gap-1"
              >
                <FileText className="w-4 h-4 text-blue-600" />
                7. 메모사항
              </label>
              <textarea
                id="taskMemo"
                name="taskMemo"
                rows={3}
                placeholder="과제에 필요한 참고사항이나 준비물을 입력하세요."
                value={memo}
                onChange={(e) => setMemo(e.target.value)}
                className="custom-input text-sm text-slate-900 placeholder:text-slate-400 resize-none"
              />
            </div>

            {/* ========================================================
                [예상 소요시간 표시 영역]
                - 스타일 요구사항: 큰 글씨(24px), 검정, 굵게, 가운데 정렬
                - 위치: 등록하기 버튼 바로 위에 배치
                - 소요시간 입력/수정 시 실시간으로 화면에 반영
                ======================================================== */}
            <div 
              aria-live="polite"
              className="mt-3 p-3.5 bg-blue-50/70 border border-blue-200/80 rounded-xl"
            >
              <p className="text-[24px] font-bold text-black text-center tracking-tight">
                예상 소요시간: {displayDuration}시간
              </p>
            </div>

            {/* [버튼 그룹: 8. 등록하기 버튼, 9. 다시 작성 버튼] */}
            <div className="grid grid-cols-2 gap-3 mt-1">
              {/* 8. 등록하기 버튼 */}
              <button
                type="submit"
                className="w-full py-3 px-4 bg-blue-600 hover:bg-blue-700 active:bg-blue-800 text-white font-bold rounded-lg shadow-sm hover:shadow transition-all duration-150 flex items-center justify-center gap-2 cursor-pointer"
              >
                <PlusCircle className="w-4 h-4" />
                등록하기
              </button>

              {/* 9. 다시 작성 버튼 */}
              <button
                type="button"
                onClick={handleReset}
                className="w-full py-3 px-4 bg-slate-100 hover:bg-slate-200 active:bg-slate-300 text-slate-700 font-bold rounded-lg border border-slate-300 shadow-sm transition-all duration-150 flex items-center justify-center gap-2 cursor-pointer"
              >
                <RotateCcw className="w-4 h-4 text-slate-500" />
                다시 작성
              </button>
            </div>

          </form>
        </section>

        {/* ========================================================
            [과제 게시판 목록 영역]
            - 등록된 과제 카드들을 동적으로 렌더링
            - 각 카드에 완료 버튼, 완료 시 배경 어두워짐 및 취소선 표시
            ======================================================== */}
        <section 
          aria-labelledby="board-heading"
          className="flex flex-col gap-4"
        >
          <div className="flex flex-wrap items-center justify-between gap-2 px-1">
            <h2 id="board-heading" className="text-lg font-bold text-slate-800 flex items-center gap-2">
              <span>📋</span> 등록된 과제 목록 
              <span className="text-xs px-2 py-0.5 rounded-full bg-blue-100 text-blue-700 font-bold">
                {taskList.length}개
              </span>
            </h2>
            <button
              type="button"
              onClick={handleDownloadExcel}
              disabled={taskList.length === 0}
              className="shrink-0 inline-flex items-center justify-center gap-1.5 rounded-lg border border-emerald-300 bg-white px-3 py-2 text-xs font-bold text-emerald-700 transition-colors hover:bg-emerald-50 disabled:cursor-not-allowed disabled:opacity-50"
            >
              <FileSpreadsheet className="h-4 w-4" />
              엑셀 다운로드
            </button>
          </div>

          {/* 등록된 과제가 없을 때의 안내 표시 */}
          {taskList.length === 0 && (
            <div className="bg-white rounded-2xl p-8 text-center border border-dashed border-slate-300 text-slate-400">
              <AlertCircle className="w-10 h-10 mx-auto mb-2 text-slate-300" />
              <p className="font-semibold text-slate-600">등록된 과제가 없습니다.</p>
              <p className="text-xs text-slate-400 mt-1">상단 폼에서 새로운 과제를 등록해보세요!</p>
            </div>
          )}

          {/* 과제 카드 리스트 동적 렌더링 */}
          <div className="flex flex-col gap-3">
            {taskList.map((task) => {
              // 중요도별 배지 색상 스타일 매핑
              const priorityColors = {
                상: 'bg-rose-100 text-rose-700 border-rose-200',
                중: 'bg-amber-100 text-amber-700 border-amber-200',
                하: 'bg-emerald-100 text-emerald-700 border-emerald-200',
              }[task.priority];

              // 분류별 뱃지 스타일
              const categoryColors: Record<string, string> = {
                개인과제: 'bg-blue-100 text-blue-700 border-blue-200',
                팀플: 'bg-purple-100 text-purple-700 border-purple-200',
                발표: 'bg-cyan-100 text-cyan-700 border-cyan-200',
              };

              return (
                <article
                  key={task.id}
                  className={`relative rounded-xl p-4 transition-all duration-200 border shadow-sm ${
                    task.isCompleted
                      ? 'bg-slate-200/90 border-slate-300 shadow-none' // 완료 시 카드 배경 어두워짐
                      : 'bg-white border-slate-200 hover:shadow-md hover:border-blue-200'
                  }`}
                >
                  {/* 카드 상단: 과제명, 삭제 버튼 */}
                  <div className="flex items-start justify-between gap-2 mb-2">
                    <h3 
                      className={`text-base font-bold text-slate-900 leading-snug break-words ${
                        task.isCompleted ? 'line-through text-slate-500' : '' // 완료 시 취소선 및 텍스트 톤 다운
                      }`}
                    >
                      {task.name}
                    </h3>
                    <button
                      type="button"
                      onClick={() => handleDeleteTask(task.id)}
                      className="text-slate-400 hover:text-rose-600 p-1 rounded-md hover:bg-rose-50 transition-colors cursor-pointer shrink-0"
                      title="과제 삭제"
                      aria-label="과제 삭제"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>

                  {/* 카드 태그 및 세부 메타 정보 */}
                  <div className="flex flex-wrap items-center gap-1.5 mb-2.5">
                    {/* 분류 뱃지 */}
                    <span className={`text-xs px-2 py-0.5 rounded-md font-semibold border ${categoryColors[task.category] || 'bg-slate-100 text-slate-700 border-slate-200'}`}>
                      {task.category}
                    </span>

                    {/* 중요도 뱃지 */}
                    <span className={`text-xs px-2 py-0.5 rounded-md font-semibold border ${priorityColors}`}>
                      중요도: {task.priority}
                    </span>

                    {/* 소요시간 뱃지 */}
                    <span className="text-xs px-2 py-0.5 rounded-md font-semibold bg-sky-50 text-sky-700 border border-sky-200 flex items-center gap-1">
                      <Clock className="w-3 h-3" />
                      {task.duration}시간 소요
                    </span>

                    {/* 날짜 표시 */}
                    {task.dueDate && (
                      <span className="text-xs px-2 py-0.5 rounded-md text-slate-600 bg-slate-100 border border-slate-200 flex items-center gap-1">
                        <Calendar className="w-3 h-3 text-slate-500" />
                        {task.dueDate}
                      </span>
                    )}
                  </div>

                  {/* 추가 옵션 태그들 (선택된 경우에만 표시) */}
                  {task.options && task.options.length > 0 && (
                    <div className="flex flex-wrap items-center gap-1 mb-2.5">
                      <span className="text-[11px] font-semibold text-slate-400 mr-0.5">옵션:</span>
                      {task.options.map((opt) => (
                        <span 
                          key={opt} 
                          className="text-[11px] px-1.5 py-0.5 rounded bg-blue-50 text-blue-600 border border-blue-100"
                        >
                          #{opt}
                        </span>
                      ))}
                    </div>
                  )}

                  {/* 메모사항 표시 (입력된 경우에만 표시) */}
                  {task.memo && (
                    <div className={`p-2.5 rounded-lg text-xs leading-relaxed mb-3 whitespace-pre-wrap ${
                      task.isCompleted ? 'bg-slate-300/50 text-slate-500' : 'bg-slate-50 text-slate-600 border border-slate-100'
                    }`}>
                      <p className="font-semibold text-slate-700 text-[11px] mb-0.5 flex items-center gap-1">
                        <FileText className="w-3 h-3" /> 메모
                      </p>
                      {task.memo}
                    </div>
                  )}

                  {/* ========================================================
                      [완료 버튼]
                      - 스타일 세부사항 준수:
                        초록 배경, 흰색 글씨, hover시 약간 밝게
                      - 클릭 시: 카드 배경 어두워짐 및 취소선 처리, 미완료 개수 실시간 차감 반영
                      ======================================================== */}
                  <div className="flex items-center justify-between pt-2 border-t border-slate-100/60 mt-2">
                    <span className="text-xs font-medium text-slate-500">
                      상태: {task.isCompleted ? (
                        <span className="text-emerald-700 font-bold">완료됨</span>
                      ) : (
                        <span className="text-amber-600 font-bold">진행 중</span>
                      )}
                    </span>

                    <button
                      type="button"
                      onClick={() => handleToggleComplete(task.id)}
                      className={`py-1.5 px-3.5 text-xs font-bold rounded-md shadow-sm transition-all duration-150 flex items-center gap-1.5 cursor-pointer ${
                        task.isCompleted
                          ? 'bg-slate-500 hover:bg-slate-600 text-white' // 완료 취소 버튼 스타일
                          : 'bg-[#16a34a] hover:bg-[#22c55e] text-white active:bg-[#15803d]' // 요구사항: 초록 배경, 흰색 글씨, hover시 약간 밝게
                      }`}
                    >
                      {task.isCompleted ? (
                        <>
                          <RotateCcw className="w-3.5 h-3.5" />
                          완료 취소
                        </>
                      ) : (
                        <>
                          <Check className="w-3.5 h-3.5 stroke-[3]" />
                          완료
                        </>
                      )}
                    </button>
                  </div>
                </article>
              );
            })}
          </div>

          {/* ========================================================
              [과제 상태 관리: 과제 미완료 확인 메시지]
              - 요구사항:
                미완료 과제가 있을 경우, 게시판 하단에 '과제 미완료 확인 메시지' 표시
                (예: "아직 완료하지 않은 과제가 2개 있습니다.")
              - 스타일 요구사항:
                회색 배경, 검정색 글씨, 둥근 모서리
              ======================================================== */}
          {uncompletedCount > 0 ? (
            <div 
              role="status"
              className="bg-gray-200 text-black rounded-lg p-3 text-center text-sm font-bold shadow-sm border border-gray-300 mt-2"
            >
              아직 완료하지 않은 과제가 {uncompletedCount}개 있습니다.
            </div>
          ) : taskList.length > 0 ? (
            <div 
              role="status"
              className="bg-emerald-100 text-emerald-900 rounded-lg p-3 text-center text-sm font-bold shadow-sm border border-emerald-200 mt-2"
            >
              🎉 대단해요! 모든 과제를 성공적으로 완료했습니다!
            </div>
          ) : null}

        </section>

        {/* 바닥글 정보 */}
        <footer className="text-center text-xs text-slate-400 py-4">
          과제 게시판 • 효율적인 하루 일정 관리
        </footer>

      </main>
    </div>
  );
}
