import React from 'react';

interface ProgressBarProps {
  progress: number;
  text: string;
  visible: boolean;
}

export const ProgressBar: React.FC<ProgressBarProps> = ({ progress, text, visible }) => {
  if (!visible) return null;
  return (
    <div className="w-full bg-gray-200 rounded-full h-2.5 mb-2 transition-all">
      <div 
        className="bg-blue-600 h-2.5 rounded-full transition-all duration-300 ease-out" 
        style={{ width: `${progress}%` }}
      ></div>
      <p className="text-xs text-center mt-1 text-gray-600 font-medium">{text}</p>
    </div>
  );
};
