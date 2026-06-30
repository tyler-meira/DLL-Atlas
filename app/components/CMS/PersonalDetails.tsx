type StatCardProps = {
  title: string;
  value: number;
  href?: string;
};

const baseStyles = "rounded-xl bg-green-500 p-6 shadow transition-colors block";

export default function StatCard({ title, value, href }: StatCardProps) {
  const CardContent = (
    <div className="rounded-xl bg-green-500 p-6 shadow">
      <>
        <p className="text-sm text-gray-500">{title}</p>
        <p className="mt-2 text-3xl font-bold">{value}</p>
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


