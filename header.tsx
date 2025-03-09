import { useState } from "react";
import Link from "next/link";
import { FiMenu, FiX } from "react-icons/fi";

const Header = () => {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <header className="bg-black text-white shadow-md">
      <div className="container mx-auto px-4 py-4 flex justify-between items-center">
        {/* Logo */}
        <Link href="/" className="text-2xl font-bold">
          StarkLotto
        </Link>

        {/* */}
        <button 
          className="md:hidden text-2xl focus:outline-none"
          onClick={() => setIsOpen(!isOpen)}
        >
          {isOpen ? <FiX /> : <FiMenu />}
        </button>

        {/* */}
        <nav className={`md:flex space-x-6 ${isOpen ? "block" : "hidden"} md:block`}>
          <Link href="/about" className="hover:text-gray-300">About</Link>
          <Link href="/games" className="hover:text-gray-300">Games</Link>
          <Link href="/contact" className="hover:text-gray-300">Contact</Link>
        </nav>
      </div>
    </header>
  );
};

export default Header;
