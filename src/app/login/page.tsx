import { redirect } from "next/navigation";
import { getSession } from "@/lib/session";

export default async function LoginPage(props: { searchParams: Promise<{ error?: string }> }) {
  const session = await getSession();
  const searchParams = await props.searchParams;
  const error = searchParams?.error;
  
  if (session.isLoggedIn) {
    redirect("/dashboard");
  }

  async function login(formData: FormData) {
    "use server";
    const username = formData.get("username") as string;
    const password = formData.get("password") as string;
    const branchId = formData.get("branchId") as string;
    const compId = formData.get("compId") as string;
    const finyear = formData.get("finyear") as string;

    if (username && password) {
      const session = await getSession();
      session.isLoggedIn = true;
      session.userId = username.toUpperCase();
      session.branchId = branchId;
      session.compId = compId;
      session.finyear = finyear;
      await session.save();
      redirect("/dashboard");
    } else {
      redirect("/login?error=invalid");
    }
  }

  return (
    <div className="flex items-center justify-center min-h-screen p-4 bg-gradient-to-br from-slate-900 via-indigo-950 to-blue-900 relative overflow-hidden bg-[size:400%_400%] animate-[gradientBG_15s_ease_infinite]">
      {/* Decorative background shapes */}
      <div className="absolute rounded-full bg-white/5 backdrop-blur-sm z-0 animate-[float_10s_infinite_ease-in-out_alternate] w-[300px] h-[300px] top-[10%] left-[15%]"></div>
      <div className="absolute rounded-full bg-white/5 backdrop-blur-sm z-0 animate-[float_10s_infinite_ease-in-out_alternate] w-[400px] h-[400px] bottom-[-10%] right-[10%] !bg-indigo-500/5 [animation-delay:-2s]"></div>
      <div className="absolute rounded-full bg-white/5 backdrop-blur-sm z-0 animate-[float_10s_infinite_ease-in-out_alternate] w-[200px] h-[200px] top-[40%] right-[25%] [animation-delay:-4s]"></div>

      <div className="w-full max-w-[420px] p-8 text-center bg-white/5 backdrop-blur-md border border-white/10 shadow-[0_25px_50px_-12px_rgba(0,0,0,0.5)] rounded-2xl z-10 relative transition-all duration-300 hover:-translate-y-1 hover:shadow-[0_30px_60px_-12px_rgba(0,0,0,0.6)]">
        <h1 className="text-3xl font-bold mb-2 text-transparent bg-clip-text bg-gradient-to-r from-indigo-100 to-indigo-300 tracking-tight">ERP System V3</h1>
        <p className="text-sm text-slate-400 mb-6">Sign in to your account</p>
        
        {error === "invalid" && (
          <div className="bg-red-500/20 border border-red-500/50 text-red-300 p-3 rounded-lg text-sm mb-6 text-left">Invalid username or password.</div>
        )}

        <form action={login} className="flex flex-col gap-4">
          <div className="flex flex-col items-start gap-2 relative">
            <label htmlFor="username" className="text-sm font-medium text-slate-300 ml-1">Username</label>
            <input 
              type="text" 
              id="username" 
              name="username" 
              placeholder="Enter your username"
              required 
              className="w-full bg-black/20 border border-white/10 rounded-xl px-4 py-3 text-white text-base transition-all duration-200 focus:outline-none focus:border-indigo-500 focus:bg-black/30 focus:ring-4 focus:ring-indigo-500/20"
            />
          </div>
          
          <div className="flex flex-col items-start gap-2 relative">
            <label htmlFor="password" className="text-sm font-medium text-slate-300 ml-1">Password</label>
            <input 
              type="password" 
              id="password" 
              name="password" 
              placeholder="Enter your password"
              required 
              className="w-full bg-black/20 border border-white/10 rounded-xl px-4 py-3 text-white text-base transition-all duration-200 focus:outline-none focus:border-indigo-500 focus:bg-black/30 focus:ring-4 focus:ring-indigo-500/20"
            />
          </div>

          <div className="grid grid-cols-2 gap-4">
            <div className="flex flex-col items-start gap-2 relative">
              <label htmlFor="branchId" className="text-sm font-medium text-slate-300 ml-1">Branch ID</label>
              <input 
                type="text" 
                id="branchId" 
                name="branchId" 
                placeholder="e.g. 00"
                required 
                className="w-full bg-black/20 border border-white/10 rounded-xl px-4 py-3 text-white text-base transition-all duration-200 focus:outline-none focus:border-indigo-500 focus:bg-black/30 focus:ring-4 focus:ring-indigo-500/20"
              />
            </div>
            
            <div className="flex flex-col items-start gap-2 relative">
              <label htmlFor="compId" className="text-sm font-medium text-slate-300 ml-1">Company ID</label>
              <input 
                type="text" 
                id="compId" 
                name="compId" 
                placeholder="e.g. 02"
                required 
                className="w-full bg-black/20 border border-white/10 rounded-xl px-4 py-3 text-white text-base transition-all duration-200 focus:outline-none focus:border-indigo-500 focus:bg-black/30 focus:ring-4 focus:ring-indigo-500/20"
              />
            </div>
          </div>

          <div className="flex flex-col items-start gap-2 relative">
            <label htmlFor="finyear" className="text-sm font-medium text-slate-300 ml-1">Financial Year</label>
            <input 
              type="text" 
              id="finyear" 
              name="finyear" 
              placeholder="e.g. 04-2026 03-2027"
              required 
              className="w-full bg-black/20 border border-white/10 rounded-xl px-4 py-3 text-white text-base transition-all duration-200 focus:outline-none focus:border-indigo-500 focus:bg-black/30 focus:ring-4 focus:ring-indigo-500/20"
            />
          </div>
          
          <button type="submit" className="mt-4 bg-gradient-to-br from-indigo-600 to-indigo-500 text-white px-4 py-3.5 rounded-xl font-semibold text-base transition-all duration-300 shadow-[0_4px_15px_rgba(79,70,229,0.4)] hover:-translate-y-0.5 hover:shadow-[0_8px_20px_rgba(79,70,229,0.6)] hover:from-indigo-700 hover:to-indigo-600 active:translate-y-px">
            Sign In
          </button>
        </form>
      </div>
    </div>
  );
}
