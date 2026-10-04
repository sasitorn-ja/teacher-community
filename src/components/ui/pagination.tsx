import type { ButtonHTMLAttributes, HTMLAttributes } from 'react'
import { ChevronLeft, ChevronRight, MoreHorizontal } from 'lucide-react'

function cn(...classes:(string|false|undefined)[]) {
  return classes.filter(Boolean).join(' ')
}

export function Pagination({className,...props}:HTMLAttributes<HTMLElement>) {
  return <nav aria-label="pagination" className={cn('pagination',className)} {...props}/>
}

export function PaginationContent({className,...props}:HTMLAttributes<HTMLUListElement>) {
  return <ul className={cn('pagination-content',className)} {...props}/>
}

export function PaginationItem({className,...props}:HTMLAttributes<HTMLLIElement>) {
  return <li className={cn('pagination-item',className)} {...props}/>
}

export function PaginationLink({className,isActive,...props}:ButtonHTMLAttributes<HTMLButtonElement> & {isActive?:boolean}) {
  return <button type="button" aria-current={isActive?'page':undefined} className={cn('pagination-link',isActive&&'is-active',className)} {...props}/>
}

export function PaginationPrevious({className,...props}:ButtonHTMLAttributes<HTMLButtonElement>) {
  return <PaginationLink aria-label="ไปหน้าก่อนหน้า" className={cn('pagination-previous',className)} {...props}><ChevronLeft size={16}/><span>ก่อนหน้า</span></PaginationLink>
}

export function PaginationNext({className,...props}:ButtonHTMLAttributes<HTMLButtonElement>) {
  return <PaginationLink aria-label="ไปหน้าถัดไป" className={cn('pagination-next',className)} {...props}><span>ถัดไป</span><ChevronRight size={16}/></PaginationLink>
}

export function PaginationEllipsis({className,...props}:HTMLAttributes<HTMLSpanElement>) {
  return <span aria-hidden="true" className={cn('pagination-ellipsis',className)} {...props}><MoreHorizontal size={16}/></span>
}
