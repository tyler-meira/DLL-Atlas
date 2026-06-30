import { LucideIcon } from "lucide-react";

type StatCardProps = {
  title: string;
  href?: string;
  icon?: LucideIcon;
};

const baseStyles = "rounded-xl bg-[#188148] p-6 shadow transition-colors block";

export default function StatCard({ title, href, icon: Icon }: StatCardProps) {
  const CardContent = (
    <div className="rounded-xl bg-white p-6 shadow flex flex-col items-center justify-center text-center">
      <>
        <p className="text-sm text-black">{title}</p>
        {Icon && (<Icon className="mt-2 w-24 h-24 object-cover rounded text-[#188148]" />)}
      </>
    </div>
  );

  if(href){
    return(
      <a href={href} target="_blank" rel="noopener noreferrer" className={`${baseStyles} cursor-pointer hover:bg-gray-100`}>
        {CardContent}
      </a>
    );
  }

  return (
    <div className={`${baseStyles}`}>
      {CardContent}
    </div>
  );

}


