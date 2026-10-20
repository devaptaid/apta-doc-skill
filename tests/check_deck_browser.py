"""Optional browser QA for the generated example; requires Playwright Chromium.

Exercises real slide navigation and print behavior, checks visible text bounds,
and captures all HTML slides for comparison with the native PowerPoint render.
"""
from __future__ import annotations
import argparse
from pathlib import Path
from playwright.sync_api import sync_playwright

ROOT=Path(__file__).resolve().parents[1]


def check(source, output):
    output.mkdir(parents=True,exist_ok=True)
    with sync_playwright() as p:
        browser=p.chromium.launch(headless=True)
        page=browser.new_page(viewport={'width':1280,'height':720})
        errors=[]
        page.on('pageerror',lambda error:errors.append(str(error)))
        page.goto(source.resolve().as_uri())
        dots=page.locator('#deck-nav .dot')
        slides=page.locator('.slide')
        count=slides.count()
        assert dots.count()==count

        def active(number):
            assert page.locator('.slide.active').count()==1
            assert slides.nth(number-1).evaluate("s=>s.classList.contains('active')")
            assert dots.nth(number-1).get_attribute('aria-current')=='true'
            assert page.locator('.dot[aria-current]').count()==1

        active(1)
        dots.nth(5).click()
        active(6)
        page.keyboard.press('ArrowRight')
        active(7)
        page.keyboard.press('Home')
        active(1)
        page.keyboard.press('End')
        active(count)
        page.keyboard.press('ArrowRight')
        active(count)
        dots.nth(5).focus()
        page.keyboard.press('Enter')
        active(6)
        page.keyboard.press('Space')
        active(6)
        page.locator('body').click(position={'x':10,'y':10})
        page.keyboard.press('Space')
        active(7)
        page.mouse.wheel(0,100)
        page.wait_for_function("document.querySelectorAll('.slide')[7].classList.contains('active')")
        active(8)
        page.mouse.wheel(0,100)
        page.wait_for_timeout(80)
        active(8)  # A single wheel gesture cannot skip through several slides.
        page.evaluate("""() => {
          const stage=document.querySelector('.deck-stage');
          const start=new Event('touchstart');start.touches=[{clientX:900,clientY:300}];stage.dispatchEvent(start);
          const end=new Event('touchend');end.changedTouches=[{clientX:600,clientY:300}];stage.dispatchEvent(end);
        }""")
        active(9)

        # Each visible text range must remain inside its actual slide canvas.
        for i in range(count):
            dots.nth(i).click()
            problems=slides.nth(i).evaluate("""s => {
              const canvas=s.getBoundingClientRect(), problems=[];
              for(const text of s.querySelectorAll('.apta-text p,th,td')){
                const range=document.createRange();range.selectNodeContents(text);
                const r=range.getBoundingClientRect();
                if(r.left<canvas.left-2||r.right>canvas.right+2||r.top<canvas.top-2||r.bottom>canvas.bottom+2)problems.push(text.textContent);
              }
              if(/editorial/i.test(s.innerText))problems.push('Unexpected Editorial label');
              return problems;
            }""")
            assert not problems,(i+1,problems)
            if slides.nth(i).get_attribute('data-layout') not in {'cover','closing'}:
                assert slides.nth(i).locator('[data-role="title"]').evaluate("e=>getComputedStyle(e).fontSize")=='48px'
            slides.nth(i).screenshot(path=str(output/f'slide-{i+1}.png'))

        for width,height in [(786,1051),(375,812)]:
            page.set_viewport_size({'width':width,'height':height})
            dots.nth(1).click()
            active(2)
            assert page.locator('.slide.active').evaluate("e=>{const r=e.getBoundingClientRect();return r.left>=-1&&r.right<=innerWidth+1&&r.top>=-1&&r.bottom<=innerHeight+1}")
            assert page.locator('#deck-nav').evaluate("e=>{const r=e.getBoundingClientRect();return r.left>=0&&r.right<=innerWidth&&r.bottom<=innerHeight}")
            page.screenshot(path=str(output/f'portrait-{width}.png'))

        page.emulate_media(media='print')
        assert not page.locator('#deck-nav').is_visible()
        assert all(slides.nth(i).is_visible() for i in range(count))
        assert not errors,errors
        browser.close()
    print(f'Browser QA passed: {count} slides; dots, keyboard, wheel, swipe, responsive canvas, print, text bounds.')


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source',type=Path,default=ROOT/'skills/apta-pptx/examples/example.html')
    parser.add_argument('--output',type=Path,default=ROOT/'.build/deck-browser-check')
    args=parser.parse_args()
    check(args.source,args.output)
