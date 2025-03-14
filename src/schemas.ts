import {string, z} from 'zod'

//
// Comment
//

export const hyperlinkSchema = z.object({
  text: z.string(),
  hyperlink: z.string().url(),
})

export const userCommentSchema = z.object({
  hyperlink: hyperlinkSchema.optional(),
  userName: z.string().min(1).optional(),
  text: z.string().optional(),
  price: z.number().nonnegative().optional(),
  numberInAlbum: z.number().nonnegative().optional(),
})

export type UserComment = z.infer<typeof userCommentSchema>

export type PurchaseType =  'RUSSIA' | 'CHINA'
