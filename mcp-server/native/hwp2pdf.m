// Headless HWP/HWPX → PDF using Hancom's own Quick Look engine (SimpleEngineLib).
// Call order taken from HwpQLExtPreview FUN_100002130: openDocument: → setZoom: → pageCount → paperExtent: → drawPageOnContext:withExtent:
//
// usage: hwp2pdf <in.hwp|in.hwpx> <out.pdf> [zoom]
// env:   HWP_APP     한/글 app bundle (default /Applications/Hancom Office HWP.app)
//        HWP_ENGINE  engine dylib (default <HWP_APP>/Contents/PlugIns/HwpQLExtPreview.appex/Contents/Frameworks/libSimpleEngineLib.dylib)
// stdout: one JSON line {"ok":true,"pages":N,"out":"..."} on success.
// exit:   0 ok · 2 usage · 3 engine load · 4 engine init · 5 open failed (corrupt/encrypted/unsupported)
//         6 engine symbols missing · 7 no pages · 8 page draw failed · 9 PDF write failed · 10 cannot take engine lock
// The PDF is written to <out>.part and renamed only after every page drew, so a crash never leaves a half PDF at <out>.
#import <Foundation/Foundation.h>
#import <AppKit/AppKit.h>
#import <objc/message.h>
#include <dlfcn.h>
#include <stdio.h>
#include <signal.h>
#include <ftw.h>
#include <sys/file.h>
#include <fcntl.h>
#include <pwd.h>
#include <unistd.h>
#include <spawn.h>
#include <crt_externs.h>
#include <mach-o/dyld.h>
#include <mach/machine.h>
typedef id (*IdFn)(id,SEL);
// Once the engine is loaded, process teardown runs Hancom C++ static destructors that abort (134) or
// segfault (139) on some documents. Our work is already on disk by then, so skip teardown.
static char g_tmp[1024];
static int rm1(const char *p,const struct stat *s,int f,struct FTW *w){(void)s;(void)f;(void)w;return remove(p);}
static void rmtmp(void){ if(g_tmp[0]) nftw(g_tmp,rm1,16,FTW_DEPTH|FTW_PHYS); }
// The engine crashes (SIGSEGV/SIGABRT) on a few documents. Remove our scratch dir, then die with the same signal.
static void oncrash(int sig){ rmtmp(); signal(sig,SIG_DFL); raise(sig); }
#define QUIT(c) do{fflush(stdout);fflush(stderr);rmtmp();_exit(c);}while(0)
int main(int argc,char**argv){@autoreleasepool{
  if(argc<3){fprintf(stderr,"usage: hwp2pdf <in.hwp|in.hwpx> <out.pdf> [zoom]\n");return 2;}
  const char *in=argv[1], *out=argv[2];
  // setZoom takes DPI: HwpQLExtPreview passes backingScale*96 and sizes pixels as extent*zoom/7200
  // (dec-ql-0x100002130.c:262-379). PDF space is 72 DPI; at 100 the layout was measured wrong (fragmented text).
  unsigned zoom = argc>3 ? (unsigned)atoi(argv[3]) : 72;
  if(zoom<36||zoom>1200){fprintf(stderr,"zoom out of range: %u\n",zoom);return 2;}
  if(access(in,R_OK)!=0){fprintf(stderr,"cannot read %s\n",in);return 5;}
  const char *app = getenv("HWP_APP") ?: "/Applications/Hancom Office HWP.app";
  // The engine links its Hancom libraries as @rpath/…. This binary carries no rpaths, so dyld resolves them
  // through DYLD_FALLBACK_LIBRARY_PATH pointed at *this* app — HWP_APP really selects the installation.
  // dyld reads DYLD_* only at launch: set it and exec ourselves once, keeping the running CPU slice so an
  // `arch -x86_64` run stays x86_64 (HWP2PDF_REEXEC guards the loop).
  if(!getenv("HWP2PDF_REEXEC")){
    char fb[2300]; const char *old=getenv("DYLD_FALLBACK_LIBRARY_PATH");
    snprintf(fb,sizeof fb,"%s/Contents/Frameworks/Hnc/Bin:%s/Contents/Frameworks%s%s",app,app,old&&*old?":":"",old&&*old?old:"");
    setenv("DYLD_FALLBACK_LIBRARY_PATH",fb,1); setenv("HWP2PDF_REEXEC","1",1);
    char self[4096]; uint32_t sz=sizeof self;
    if(_NSGetExecutablePath(self,&sz)!=0){fprintf(stderr,"cannot find own executable\n");return 3;}
    posix_spawnattr_t at; posix_spawnattr_init(&at);
    posix_spawnattr_setflags(&at,POSIX_SPAWN_SETEXEC);
#if defined(__x86_64__)
    cpu_type_t cpu=CPU_TYPE_X86_64;
#else
    cpu_type_t cpu=CPU_TYPE_ARM64;
#endif
    size_t ok=0; posix_spawnattr_setbinpref_np(&at,1,&cpu,&ok);
    posix_spawn(NULL,self,NULL,&at,argv,*_NSGetEnviron());
    fprintf(stderr,"re-exec failed\n"); return 3;
  }
  unsetenv("HWP2PDF_REEXEC");

  // The engine's compound-file layer (libhsp) puts scratch files at $TMPDIR/hspcf.XXXXXX. Give every
  // process its own TMPDIR so concurrent runs share nothing on disk.
  char tmpl[]="/tmp/hwp2pdf.XXXXXX";
  const char *base=getenv("TMPDIR");
  char tmpdir[1024]; snprintf(tmpdir,sizeof tmpdir,"%s/hwp2pdf.XXXXXX",base&&*base?base:"/tmp");
  char *own=mkdtemp(tmpdir); if(!own) own=mkdtemp(tmpl);
  if(own){ setenv("TMPDIR",own,1); snprintf(g_tmp,sizeof g_tmp,"%s",own); }
  int sigs[]={SIGSEGV,SIGBUS,SIGABRT,SIGILL,SIGFPE,SIGTERM,SIGINT};
  for(unsigned k=0;k<sizeof sigs/sizeof *sigs;k++) signal(sigs[k],oncrash);
  [NSApplication sharedApplication];
  char engine[1024];
  if(getenv("HWP_ENGINE")) snprintf(engine,sizeof engine,"%s",getenv("HWP_ENGINE"));
  else snprintf(engine,sizeof engine,"%s/Contents/PlugIns/HwpQLExtPreview.appex/Contents/Frameworks/libSimpleEngineLib.dylib",app);
  if(!dlopen(engine,RTLD_NOW)){fprintf(stderr,"dlopen: %s\n",dlerror());QUIT(3);}
  Class C=NSClassFromString(@"SimpleEngineLib");
  if(!C){fprintf(stderr,"SimpleEngineLib class missing\n");QUIT(6);}
  id e=((IdFn)objc_msgSend)([C alloc],@selector(init));
  if(!e){fprintf(stderr,"init nil\n");QUIT(4);}
  // initHncDirs derives these from the host app bundle; we are not inside 한/글, so set them directly
  // (HncGlobalSettings::SetString, same keys initHncDirs writes — dec-se-0x1a9c.c).
  void (*setStr)(const char*,const char*) = dlsym(RTLD_DEFAULT,"_ZN17HncGlobalSettings9SetStringEPKcS1_");
  void (*hspSet)(const char*,const char*) = dlsym(RTLD_DEFAULT,"_Z26HspGlobalSettingsSetStringPKcS0_");
  if(!setStr){fprintf(stderr,"no HncGlobalSettings::SetString\n");QUIT(6);}
  char res[1024], bin[1024];
  snprintf(res,sizeof res,"%s/Contents/Resources/Hnc",app);
  snprintf(bin,sizeof bin,"%s/Contents/Frameworks/Hnc/Bin",app);
  setStr("HNC_DIR",res); setStr("HNC_BIN_DIR",bin);
  if(hspSet) hspSet("LOCALE_LANG","ko");
  // libhsp's temp dir defaults to NSTemporaryDirectory(), shared with every app of this user, and
  // HspTempClearDir empties it. Point it at our private scratch dir before the engine creates anything.
  void (*hspTempSetDir)(const char*) = dlsym(RTLD_DEFAULT,"_Z13HspTempSetDirPKc");
  if(!hspTempSetDir){fprintf(stderr,"no HspTempSetDir\n");QUIT(6);}
  if(g_tmp[0]) hspTempSetDir(g_tmp);

  // Several engines drawing at once sometimes render a page blank while every call still reports success
  // (measured: 3/180 blank PDFs with 12 concurrent runs, 0/180 when serialized). Serialize the engine work
  // with a per-user flock; the kernel drops it when the process exits, crash included. The lock lives at a
  // fixed per-user path, not under TMPDIR: callers with different TMPDIRs must still share one lock.
  // HOME may be unset (launchd, some MCP hosts); the passwd entry gives the same directory, so every caller of
  // this user still meets on one lock file. If the lock cannot be taken, do not convert at all.
  const char *home=getenv("HOME");
  if(!home||!*home){ struct passwd *pw=getpwuid(getuid()); home=pw&&pw->pw_dir?pw->pw_dir:NULL; }
  if(!home){fprintf(stderr,"lock: no home directory\n");QUIT(10);}
  char lockp[1100]; snprintf(lockp,sizeof lockp,"%s/Library/Caches/hwp2pdf.lock",home);
  fprintf(stderr,"lock: %s\n",lockp);
  int lockfd=open(lockp,O_CREAT|O_RDWR|O_CLOEXEC,0600);
  if(lockfd<0||flock(lockfd,LOCK_EX)!=0){perror("lock");QUIT(10);}
  // openDocument: fails intermittently when several processes start at once (measured: 15/80 with 8
  // concurrent runs, 0/80 with up to 3 attempts; 0/10 sequential). Retry only this step — an encrypted or
  // corrupt file fails every attempt, so its result is unchanged.
  NSString *inPath=[NSString stringWithUTF8String:in];
  BOOL ok=NO;
  for(int a=0;a<3&&!ok;a++){
    if(a) usleep(150000*a+(arc4random_uniform(200000)));
    ok=((BOOL(*)(id,SEL,id))objc_msgSend)(e,@selector(openDocument:),inPath);
  }
  if(!ok){fprintf(stderr,"openDocument failed\n");QUIT(5);}
  ((void(*)(id,SEL,unsigned))objc_msgSend)(e,@selector(setZoom:),zoom);
  unsigned n=((unsigned(*)(id,SEL))objc_msgSend)(e,@selector(pageCount));
  fprintf(stderr,"pages=%u\n",n);
  if(n==0){fprintf(stderr,"no pages\n");QUIT(7);}

  char part[1100]; snprintf(part,sizeof part,"%s.part",out);
  unlink(part);
  NSURL *partURL=[NSURL fileURLWithPath:[NSString stringWithUTF8String:part]];
  NSDictionary *info=@{(__bridge NSString*)kCGPDFContextCreator:@"hwp2pdf (Hancom SimpleEngine)",
                       (__bridge NSString*)kCGPDFContextTitle:[[NSString stringWithUTF8String:in] lastPathComponent]};
  CGContextRef pdf=NULL;
  unsigned failed=0;
  for(unsigned i=0;i<n;i++){
    // CGRect is 32 bytes: arm64 returns it in registers, x86_64 through a hidden pointer (objc_msgSend_stret).
#if defined(__x86_64__)
    CGRect r=((CGRect(*)(id,SEL,unsigned))objc_msgSend_stret)(e,@selector(paperExtent:),i);
#else
    CGRect r=((CGRect(*)(id,SEL,unsigned))objc_msgSend)(e,@selector(paperExtent:),i);
#endif
    // paperExtent is in HWPUNIT (1/7200 in); PDF points are 1/72 in.
    CGRect box=CGRectMake(0,0,r.size.width/100.0,r.size.height/100.0);
    fprintf(stderr,"page %u extent=%.0fx%.0f -> %.1fx%.1fpt\n",i,r.size.width,r.size.height,box.size.width,box.size.height);
    if(box.size.width<1||box.size.height<1){fprintf(stderr,"page %u has empty extent\n",i);failed++;continue;}
    if(!pdf){pdf=CGPDFContextCreateWithURL((__bridge CFURLRef)partURL,&box,(__bridge CFDictionaryRef)info);
      if(!pdf){fprintf(stderr,"cannot create %s\n",part);QUIT(9);}}
    CGContextBeginPage(pdf,&box);
    CGContextSaveGState(pdf);
    CGContextTranslateCTM(pdf,0,box.size.height); CGContextScaleCTM(pdf,1,-1);
    BOOL d=((BOOL(*)(id,SEL,CGContextRef,CGRect))objc_msgSend)(e,@selector(drawPageOnContext:withExtent:),pdf,r);
    CGContextRestoreGState(pdf);
    CGContextEndPage(pdf);
    if(!d){fprintf(stderr,"draw %u failed\n",i);failed++;}
  }
  if(pdf){CGPDFContextClose(pdf);CGContextRelease(pdf);}
  // Never call releaseEngine. It ends in libhsp HspTempClearDir, which recursively empties the engine's
  // temp dir. That dir defaults to NSTemporaryDirectory() — the user's whole /var/folders/../T, shared with
  // every other app (REA: HspTempGetDir → HspAppGetTemporaryDirectory → NSTemporaryDirectory; trace by
  // interposing rmdir/unlink showed one run removing com.apple.*, AppTranslocation, xcrun_db there).
  // HspTempSetDir above already confines it to our scratch dir; skipping release is the second guard.
  // We _exit right after, and rmtmp() removes the scratch dir with everything the engine left in it.
  if(failed){unlink(part);fprintf(stderr,"%u/%u pages failed\n",failed,n);QUIT(8);}
  if(rename(part,out)!=0){perror("rename");unlink(part);QUIT(9);}
  NSString *o=[NSString stringWithUTF8String:out];
  NSData *j=[NSJSONSerialization dataWithJSONObject:@{@"ok":@YES,@"pages":@(n),@"out":o} options:0 error:nil];
  fwrite(j.bytes,1,j.length,stdout); fputc('\n',stdout);
  QUIT(0);
}}
