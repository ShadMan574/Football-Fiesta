'use client';

import React, { useState } from 'react';
import { useApp } from '@/context/AppContext';
import { PlayerCard } from '@/components/PlayerCard';
import { Player, PlayerPosition, PlayerStatus } from '@/types';
import * as XLSX from 'xlsx';
import { 
  Users, 
  Search, 
  Filter, 
  FileSpreadsheet, 
  Plus, 
  Upload, 
  Lock, 
  UserPlus,
  Trash2,
  X
} from 'lucide-react';

const getGoogleDriveFileId = (value: string) => {
  const match = value.match(/(?:\/file\/d\/|[?&]id=|\/uc\?id=)([a-zA-Z0-9_-]+)/);
  return match?.[1] ?? null;
};

const normalizeImageUrl = (value: unknown) => {
  if (typeof value !== 'string') return undefined;

  const url = value.trim();
  if (!url) return undefined;

  const fileId = getGoogleDriveFileId(url);
  return fileId
    ? `https://drive.google.com/thumbnail?id=${fileId}&sz=w1000`
    : url;
};

export default function PlayersDirectoryPage() {
  const { isAdmin, players, addPlayer, deletePlayer, bulkImportPlayers } = useApp();

  const [searchTerm, setSearchTerm] = useState('');
  const [positionFilter, setPositionFilter] = useState<string>('ALL');
  const [statusFilter, setStatusFilter] = useState<string>('ALL');

  // Add single player modal state
  const [isAddModalOpen, setIsAddModalOpen] = useState(false);
  const [newName, setNewName] = useState('');
  const [newRoll, setNewRoll] = useState('');
  const [newSeries, setNewSeries] = useState('21 Series');
  const [newPosition, setNewPosition] = useState<PlayerPosition>('FORWARD');

  // Bulk import state
  const [isImportModalOpen, setIsImportModalOpen] = useState(false);
  const [importReplace, setImportReplace] = useState(false);
  const [importCount, setImportCount] = useState<number | null>(null);

  // Filtered Players list
  const filteredPlayers = players.filter((p) => {
    const matchesSearch = 
      p.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
      p.roll.toLowerCase().includes(searchTerm.toLowerCase()) ||
      p.series.toLowerCase().includes(searchTerm.toLowerCase()) ||
      (p.teamName && p.teamName.toLowerCase().includes(searchTerm.toLowerCase()));

    const matchesPosition = positionFilter === 'ALL' || p.position === positionFilter;
    const matchesStatus = statusFilter === 'ALL' || p.status === statusFilter;

    return matchesSearch && matchesPosition && matchesStatus;
  });

  // Handle Add Single Player
  const handleAddPlayer = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newName || !newRoll) return;

    addPlayer({
      name: newName,
      roll: newRoll,
      series: newSeries,
      position: newPosition,
      isIcon: false
    });

    setNewName('');
    setNewRoll('');
    setIsAddModalOpen(false);
  };

  // Handle Excel / CSV Import
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json<any>(ws);

        const imported: Player[] = data.map((row, idx) => ({
          id: 'imp-' + Date.now() + '-' + idx,
          name: row.Name || row.name || `Player ${idx + 1}`,
          roll: String(row.Roll || row.roll || `2003${idx + 10}`),
          series: row.Series || row.series || '21 Series',
          position: (row.Position || row.position || 'FORWARD').toUpperCase() as PlayerPosition,
          isIcon: false,
          status: 'AVAILABLE' as PlayerStatus,
          goalsScored: 0,
          photoUrl: normalizeImageUrl(
            row['Photo URL'] || row['Photo Url'] || row.photoUrl || row.Photo ||
            row.photo || row['Image URL'] || row['Image Url'] || row.imageUrl || row.Image || row.image
          )
        }));

        if (imported.length > 0) {
          bulkImportPlayers(imported, importReplace);
          setImportCount(imported.length);
          setTimeout(() => {
            setIsImportModalOpen(false);
            setImportCount(null);
          }, 1500);
        }
      } catch (err) {
        alert('Failed to parse Excel/CSV file. Please ensure columns include: Name, Roll, Series, Position.');
      }
    };
    reader.readAsBinaryString(file);
  };

  return (
    <div className="visual-rally rally-players max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-8">
      
      {/* Header Banner */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-gray-800 pb-6">
        <div className="flex items-center gap-3">
          <div className="p-3 bg-gradient-to-r from-teal to-light-cyan rounded-xl text-charcoal shadow-glow-cyan">
            <Users className="w-8 h-8" />
          </div>
          <div>
            <h1 className="font-bebas text-5xl text-white tracking-wide">PLAYER DIRECTORY & POOL</h1>
            <p className="text-xs text-light-cyan font-montserrat">Searchable registry of registered players with CSV/Excel bulk import</p>
          </div>
        </div>

        {isAdmin ? (
          <div className="flex items-center gap-3">
            <button
              onClick={() => setIsImportModalOpen(true)}
              className="px-4 py-2.5 bg-deep-blue hover:bg-gray-800 text-light-cyan font-bebas text-lg rounded-xl border border-light-cyan/40 hover:border-light-cyan transition-all flex items-center gap-2"
            >
              <FileSpreadsheet size={18} />
              <span>Bulk Excel/CSV Import</span>
            </button>

            <button
              onClick={() => setIsAddModalOpen(true)}
              className="px-4 py-2.5 bg-primary-yellow text-charcoal font-bebas text-lg font-bold rounded-xl shadow-glow-yellow hover:opacity-90 transition-all flex items-center gap-2"
            >
              <UserPlus size={18} />
              <span>Add Player</span>
            </button>
          </div>
        ) : (
          <div className="flex items-center gap-2 text-xs text-gray-400 bg-charcoal/80 px-4 py-2 rounded-xl border border-gray-700">
            <Lock size={16} className="text-primary-yellow" />
            <span>Admin login required for import/add</span>
          </div>
        )}
      </div>

      {/* Filters & Search Toolbar */}
      <div className="glass-panel p-6 rounded-2xl space-y-4">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          
          {/* Search Box */}
          <div className="relative">
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              placeholder="Search by name, roll, series, or team..."
              className="w-full bg-charcoal/90 border border-gray-700 rounded-xl px-4 py-2.5 pl-10 text-sm text-white placeholder-gray-500 focus:border-primary-yellow focus:outline-none"
            />
            <Search className="absolute left-3 top-3 text-gray-400 w-4 h-4" />
          </div>

          {/* Position Filter */}
          <div className="flex items-center gap-2">
            <Filter size={16} className="text-light-cyan shrink-0" />
            <select
              value={positionFilter}
              onChange={(e) => setPositionFilter(e.target.value)}
              className="w-full bg-charcoal/90 border border-gray-700 rounded-xl px-4 py-2.5 text-xs text-white focus:border-primary-yellow focus:outline-none"
            >
              <option value="ALL">All Positions</option>
              <option value="FORWARD">FORWARD</option>
              <option value="MIDFIELDER">MIDFIELDER</option>
              <option value="DEFENDER">DEFENDER</option>
              <option value="GOALKEEPER">GOALKEEPER</option>
            </select>
          </div>

          {/* Status Filter */}
          <div className="flex items-center gap-2">
            <select
              value={statusFilter}
              onChange={(e) => setStatusFilter(e.target.value)}
              className="w-full bg-charcoal/90 border border-gray-700 rounded-xl px-4 py-2.5 text-xs text-white focus:border-primary-yellow focus:outline-none"
            >
              <option value="ALL">All Statuses</option>
              <option value="AVAILABLE">AVAILABLE (Pool)</option>
              <option value="SOLD">SOLD</option>
              <option value="UNSOLD">UNSOLD</option>
              <option value="ICON">ICON PLAYERS</option>
            </select>
          </div>

        </div>
      </div>

      {/* Grid of Players */}
      <div>
        <div className="flex items-center justify-between mb-4">
          <p className="text-xs text-gray-400 font-montserrat">
            Showing <strong className="text-primary-yellow">{filteredPlayers.length}</strong> of {players.length} players
          </p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {filteredPlayers.map((player) => (
            <div key={player.id} className="relative group">
              <PlayerCard player={player} />
              
              {isAdmin && (
                <button
                  onClick={() => deletePlayer(player.id)}
                  className="absolute top-3 right-3 opacity-0 group-hover:opacity-100 p-2 bg-fiery-red text-white rounded-lg transition-opacity hover:scale-110 shadow-md z-20"
                  title="Delete Player"
                >
                  <Trash2 size={14} />
                </button>
              )}
            </div>
          ))}
        </div>
      </div>

      {/* Add Single Player Modal */}
      {isAddModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-md p-4">
          <div className="relative w-full max-w-md glass-panel rounded-2xl p-6 text-white space-y-4">
            <button
              onClick={() => setIsAddModalOpen(false)}
              className="absolute top-4 right-4 text-gray-400 hover:text-white"
            >
              <X size={20} />
            </button>

            <h3 className="font-bebas text-3xl text-primary-yellow">Register New Player</h3>

            <form onSubmit={handleAddPlayer} className="space-y-4">
              <div>
                <label className="block text-xs font-semibold text-gray-300 uppercase mb-1">Full Name</label>
                <input
                  type="text"
                  value={newName}
                  onChange={(e) => setNewName(e.target.value)}
                  placeholder="e.g. Mahfuz Ahmed"
                  className="w-full bg-deep-blue border border-gray-700 rounded-xl px-4 py-2.5 text-sm text-white focus:border-primary-yellow focus:outline-none"
                  required
                />
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-semibold text-gray-300 uppercase mb-1">Roll No.</label>
                  <input
                    type="text"
                    value={newRoll}
                    onChange={(e) => setNewRoll(e.target.value)}
                    placeholder="e.g. 2103045"
                    className="w-full bg-deep-blue border border-gray-700 rounded-xl px-4 py-2.5 text-sm text-white focus:border-primary-yellow focus:outline-none"
                    required
                  />
                </div>

                <div>
                  <label className="block text-xs font-semibold text-gray-300 uppercase mb-1">Series</label>
                  <select
                    value={newSeries}
                    onChange={(e) => setNewSeries(e.target.value)}
                    className="w-full bg-deep-blue border border-gray-700 rounded-xl px-4 py-2.5 text-xs text-white focus:border-primary-yellow focus:outline-none"
                  >
                    <option value="19 Series">19 Series</option>
                    <option value="20 Series">20 Series</option>
                    <option value="21 Series">21 Series</option>
                    <option value="22 Series">22 Series</option>
                    <option value="23 Series">23 Series</option>
                  </select>
                </div>
              </div>

              <div>
                <label className="block text-xs font-semibold text-gray-300 uppercase mb-1">Position</label>
                <select
                  value={newPosition}
                  onChange={(e) => setNewPosition(e.target.value as PlayerPosition)}
                  className="w-full bg-deep-blue border border-gray-700 rounded-xl px-4 py-2.5 text-xs text-white focus:border-primary-yellow focus:outline-none"
                >
                  <option value="FORWARD">FORWARD</option>
                  <option value="MIDFIELDER">MIDFIELDER</option>
                  <option value="DEFENDER">DEFENDER</option>
                  <option value="GOALKEEPER">GOALKEEPER</option>
                </select>
              </div>

              <div className="flex gap-3 pt-2">
                <button
                  type="button"
                  onClick={() => setIsAddModalOpen(false)}
                  className="flex-1 py-2.5 bg-gray-800 text-gray-300 font-bebas text-lg rounded-xl"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="flex-1 py-2.5 bg-primary-yellow text-charcoal font-bebas text-xl font-bold rounded-xl shadow-glow-yellow"
                >
                  Save Player
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Bulk Excel/CSV Import Modal */}
      {isImportModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-md p-4">
          <div className="relative w-full max-w-lg glass-panel-gold rounded-3xl p-8 text-white space-y-6">
            <button
              onClick={() => setIsImportModalOpen(false)}
              className="absolute top-4 right-4 text-gray-400 hover:text-white"
            >
              <X size={20} />
            </button>

            <div className="flex items-center gap-3 border-b border-primary-yellow/30 pb-4">
              <FileSpreadsheet className="w-8 h-8 text-primary-yellow" />
              <div>
                <h3 className="font-bebas text-3xl text-primary-yellow">Bulk Excel / CSV Import</h3>
                <p className="text-xs text-gray-300">Upload .xlsx or .csv spreadsheet file to populate player pool</p>
              </div>
            </div>

            {importCount !== null ? (
              <div className="text-center py-6 space-y-2 bg-teal/10 border border-teal/30 rounded-xl">
                <p className="font-bebas text-2xl text-teal">Successfully Imported {importCount} Players!</p>
              </div>
            ) : (
              <div className="space-y-4">
                <div className="flex items-center gap-2">
                  <input
                    type="checkbox"
                    id="importReplace"
                    checked={importReplace}
                    onChange={(e) => setImportReplace(e.target.checked)}
                    className="w-4 h-4 accent-primary-yellow"
                  />
                  <label htmlFor="importReplace" className="text-xs text-gray-300">
                    Replace existing auction pool players (preserves Icon players)
                  </label>
                </div>

                <div className="border-2 border-dashed border-primary-yellow/40 rounded-2xl p-8 text-center bg-charcoal/80 hover:bg-deep-blue/80 transition-colors">
                  <Upload className="w-10 h-10 text-primary-yellow mx-auto mb-2 animate-pulse" />
                  <p className="text-sm font-semibold text-white mb-1">Click or drag Excel / CSV file here</p>
                  <p className="text-xs text-gray-400 mb-4">Supported formats: .xlsx, .xls, .csv</p>
                  
                  <input
                    type="file"
                    accept=".xlsx, .xls, .csv"
                    onChange={handleFileUpload}
                    className="hidden"
                    id="excelFileInput"
                  />
                  <label
                    htmlFor="excelFileInput"
                    className="inline-block px-6 py-2.5 bg-primary-yellow text-charcoal font-bebas text-lg font-bold rounded-xl cursor-pointer shadow-glow-yellow"
                  >
                    Select Spreadsheet File
                  </label>
                </div>

                <div className="text-[11px] text-gray-400 bg-charcoal/80 p-3 rounded-xl border border-gray-800 space-y-1">
                  <p className="font-semibold text-gray-200">Required Column Headers:</p>
                  <p className="font-mono text-primary-yellow">Name | Roll | Series | Position | Photo URL</p>
                  <p>For Google Drive images, set General access to <strong className="text-gray-200">Anyone with the link</strong> and paste the file link.</p>
                </div>
              </div>
            )}
          </div>
        </div>
      )}

    </div>
  );
}
